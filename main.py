import sys
import calendar
import pandas as pd
from PyQt5.QtWidgets import QComboBox, QFileDialog
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox, QFrame, QDialog,
    QDateEdit, QLineEdit, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont
from supabase import create_client, Client

# ========================
# KONFIGURASI SUPABASE
# ========================
SUPABASE_URL = "https://rimeaufssicbkjbbwgul.supabase.co"  # Ganti
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJpbWVhdWZzc2ljYmtqYmJ3Z3VsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjEwNzk2NjcsImV4cCI6MjA3NjY1NTY2N30.-rZuJajqEwbf0F9hWC-iPIFf0lDtCGEcm8Zwt6FdXv8"  # Ganti

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# ====================================================
# KELAS UTAMA APLIKASI
# ====================================================
class PembukuanApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Pembukuan Sederhana")
        self.resize(1100, 650)
        self.setFont(QFont("Segoe UI", 10))
        self.saldo_sebelumnya = 0
        self.initUI()

    def initUI(self):
        main_layout = QHBoxLayout(self)
        
        # ---------------- Sidebar ----------------
        sidebar = QVBoxLayout()
        sidebar.setSpacing(10)
        
        sidebar_frame = QFrame()
        sidebar_frame.setStyleSheet("""
            QFrame {
                background-color: #B3E5FC;
                border-radius: 5px;
            }
            QPushButton {
                background-color: #0288D1;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                text-align: left;
            }
            QPushButton:hover {
                background-color: #0277BD;
            }
        """)
        
        self.btn_tambah = QPushButton("➕ Tambah Data")
        self.btn_edit = QPushButton("✏️ Edit Data")
        self.btn_hapus = QPushButton("🗑️ Hapus Data")
        self.btn_lihat_semua = QPushButton("📋 Lihat Semua Data")
        self.btn_periodik = QPushButton("📅 Lihat Data Periodik")
        self.btn_tahunan = QPushButton("📆 Lihat Data Tahunan")
        # ... tambah sidebar ke layout utama
        for btn in [self.btn_tambah, self.btn_edit, self.btn_hapus,
                    self.btn_lihat_semua, self.btn_periodik, self.btn_tahunan]:
            sidebar.addWidget(btn)

        sidebar.addStretch()
        sidebar_frame.setLayout(sidebar)
        main_layout.addWidget(sidebar_frame, 1)
        # --- Tambahkan tombol Exit di bawah tombol lain ---
        sidebar.addStretch()
        self.btn_exit = QPushButton("❌ Keluar")
        self.btn_exit.setStyleSheet("""
            QPushButton {
                background-color: black;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                text-align: left;
            }
            QPushButton:hover {
                background-color: #333333;
            }
        """)
        # === Tambahkan logo dan keterangan di bawah sidebar ===
        from PyQt5.QtGui import QPixmap
        sidebar.addStretch(2) 
        logo_bottom = QLabel()
        logo_bottom.setPixmap(QPixmap("logo.png").scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        logo_bottom.setAlignment(Qt.AlignCenter)
        lbl_ket = QLabel("AsratChannel 2025")
        lbl_ket.setAlignment(Qt.AlignCenter)
        lbl_ket.setStyleSheet("""
            QLabel {
                color: #01579B;
                font-weight: bold;
                font-size: 10pt;
            }
        """)
        sidebar.addWidget(logo_bottom)
        sidebar.addWidget(lbl_ket)
        sidebar.addSpacing(10)

        # Tombol Exit tetap paling bawah
        sidebar.addWidget(self.btn_exit)
        self.btn_exit.clicked.connect(QApplication.instance().quit)
        # ---------------- Konten ----------------
        content_layout = QVBoxLayout()
        self.lbl_title = QLabel("📘 Data Pembukuan Arus Kas Harian")
        self.lbl_title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        content_layout.addWidget(self.lbl_title)

        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)  # aktifkan warna baris selang-seling
        self.table.setStyleSheet("""
            QTableWidget {
                gridline-color: #B0BEC5;  /* warna garis antar kolom/baris */
                background-color: #FAFAFA;
                alternate-background-color: #E3F2FD;  /* warna baris selang-seling */
                selection-background-color: #BBDEFB;  /* warna saat dipilih */
                selection-color: #000000;  /* warna teks saat dipilih */
                border: 1px solid #90A4AE;
            }
            QHeaderView::section {
                background-color: #0288D1;  /* warna header kolom */
                color: white;
                padding: 6px;
                font-weight: bold;
                border: 1px solid #90A4AE;
            }
            QTableWidget::item {
                padding: 4px;
            }
        """)
        self.table.setShowGrid(True)  # tampilkan grid antar kolom/baris
        self.table.setSelectionBehavior(QTableWidget.SelectRows)  # pilih seluruh baris
        self.table.setSelectionMode(QTableWidget.SingleSelection)  # hanya satu baris sekaligus
        content_layout.addWidget(self.table)

        self.lbl_saldo_sebelumnya = QLabel("💼 Saldo Kas Sebelumnya = Rp 0.00")
        self.lbl_saldo_sebelumnya.setStyleSheet("""
            QLabel {
                background-color: #FFF9C4;      /* kuning lembut */
                border: 1px solid #FDD835;
                border-radius: 5px;
                padding: 8px;
                font-size: 13pt;
                font-weight: bold;
                color: #424242;
            }
        """)

        self.lbl_ringkasan = QLabel("💰 Kas Masuk = Rp0 | Kas Keluar = Rp0 | Saldo Akhir = Rp0")
        self.lbl_ringkasan.setStyleSheet("""
            QLabel {
                background-color: #C8E6C9;      /* hijau lembut */
                border: 1px solid #81C784;
                border-radius: 5px;
                padding: 8px;
                font-size: 13pt;
                font-weight: bold;
                color: #1B5E20;
            }
        """)

        content_layout.addWidget(self.lbl_saldo_sebelumnya)
        content_layout.addWidget(self.lbl_ringkasan)

        main_layout.addLayout(content_layout, 4)

        # Tombol Event
        self.btn_lihat_semua.clicked.connect(self.load_data)
        self.btn_tambah.clicked.connect(self.tambah_data)
        self.btn_edit.clicked.connect(self.edit_data)
        self.btn_hapus.clicked.connect(self.hapus_data)
        self.btn_periodik.clicked.connect(self.lihat_periodik)
        self.btn_tahunan.clicked.connect(self.lihat_tahunan)

        self.load_data()

    def format_rupiah(self, nilai):
    #"""Format angka menjadi format Rupiah: 1234567.8 -> 1.234.567,80"""
        try:
            return f"{float(nilai):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(nilai)
    
    # ====================================================
    # MUAT DATA NORMAL (default / periodik)
    # ====================================================
    def load_data(self, start_date=None, end_date=None, year=None, saldo_sebelumnya=None):
        try:
                        # Reset judul setiap kali load data normal
            if start_date and end_date:
                self.lbl_title.setText(f"📅 Laporan Arus Kas Periodik {start_date} s.d {end_date}")
            elif year:
                self.lbl_title.setText(f"📆 Rekap Arus Kas Tahunan {year}")
            else:
                self.lbl_title.setText("📆 Data Transaksi Arus Kas")
            query = supabase.table("pembukuan").select("*")

            if start_date and end_date:
                query = query.gte("tanggal", start_date).lte("tanggal", end_date)
            elif year:
                query = query.like("tanggal", f"{year}%")

            data = query.order("tanggal", desc=False).execute().data
            self.table.setColumnCount(7)
            self.table.setHorizontalHeaderLabels(["Nomor", "No.Transaksi", "Tgl.Transaksi", "Deskripsi", "Uang Masuk", "Uang Keluar", "Total Uang"])
            self.table.setRowCount(len(data))

            total_masuk = total_keluar = saldo_akhir = 0

            for row, item in enumerate(data):
                self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))  # Kolom Nomor (indeks 1,2,...)
                self.table.setItem(row, 1, QTableWidgetItem(str(item['id'])))
                self.table.setItem(row, 2, QTableWidgetItem(str(item['tanggal'])))
                self.table.setItem(row, 3, QTableWidgetItem(item['keterangan']))
                # Format ke rupiah dengan titik ribuan dan koma desimal
                def format_rupiah(nilai):
                    try:
                        return f"{float(nilai):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    except:
                        return str(nilai)
                #---membuat format rupiah dan rata kanan ------
                item_debet = QTableWidgetItem(self.format_rupiah(item['debet']))
                item_debet.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, 4, item_debet)

                item_kredit = QTableWidgetItem(self.format_rupiah(item['kredit']))
                item_kredit.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, 5, item_kredit)

                item_saldo = QTableWidgetItem(self.format_rupiah(item['total_saldo']))
                item_saldo.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, 6, item_saldo)

                total_masuk += float(item['debet'])
                total_keluar += float(item['kredit'])
                saldo_akhir = float(item['total_saldo'])

            if saldo_sebelumnya is not None:
                self.saldo_sebelumnya = saldo_sebelumnya

            self.lbl_saldo_sebelumnya.setText(f"💼 Saldo Kas Sebelumnya = Rp {self.saldo_sebelumnya:,.2f}")
            self.lbl_ringkasan.setText(
                f"💰 Total Kas Masuk = Rp{total_masuk:,.2f} 💰 Total Kas Keluar = Rp{total_keluar:,.2f} 💰 Total Saldo Akhir = Rp{saldo_akhir:,.2f}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal memuat data: {e}")

    # ====================================================
    # TAMBAH DATA (DENGAN INPUT FORM MANUAL)
    # ====================================================
    def tambah_data(self):
        # --- Buat dialog input ---
        dialog = QDialog(self)
        dialog.setWindowTitle("Tambah Data Transaksi")
        layout = QVBoxLayout(dialog)

        # --- Input field ---
        lbl_tanggal = QLabel("Tanggal Transaksi:")
        date_edit = QDateEdit(QDate.currentDate())
        date_edit.setCalendarPopup(True)

        lbl_keterangan = QLabel("Keterangan:")
        txt_keterangan = QLineEdit()
        txt_keterangan.setPlaceholderText("Contoh: Pembelian alat tulis")

        lbl_debet = QLabel("Uang Masuk (Rp):")
        txt_debet = QLineEdit()
        txt_debet.setPlaceholderText("Isi 0 jika tidak ada")

        lbl_kredit = QLabel("Uang Keluar (Rp):")
        txt_kredit = QLineEdit()
        txt_kredit.setPlaceholderText("Isi 0 jika tidak ada")

        layout.addWidget(lbl_tanggal)
        layout.addWidget(date_edit)
        layout.addWidget(lbl_keterangan)
        layout.addWidget(txt_keterangan)
        layout.addWidget(lbl_debet)
        layout.addWidget(txt_debet)
        layout.addWidget(lbl_kredit)
        layout.addWidget(txt_kredit)

        # Tombol OK / Cancel
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)

        # --- Tampilkan dialog ---
        if dialog.exec_() == QDialog.Accepted:
            tanggal = date_edit.date().toString("yyyy-MM-dd")
            keterangan = txt_keterangan.text().strip()
            debet = txt_debet.text().strip()
            kredit = txt_kredit.text().strip()

            # Validasi input
            try:
                debet = float(debet) if debet else 0.0
                kredit = float(kredit) if kredit else 0.0
            except ValueError:
                QMessageBox.warning(self, "Peringatan", "Masukkan angka yang valid untuk Uang Masuk / Keluar!")
                return

            if not keterangan:
                QMessageBox.warning(self, "Peringatan", "Keterangan harus diisi!")
                return

            try:
                # Ambil saldo terakhir
                res = supabase.table("pembukuan") \
                    .select("total_saldo") \
                    .order("tanggal", desc=True) \
                    .limit(1) \
                    .execute().data
                saldo_terakhir = float(res[0]['total_saldo']) if res else 0

                # Hitung saldo baru
                total_saldo = saldo_terakhir + (debet - kredit)

                # Simpan ke database
                supabase.table("pembukuan").insert({
                    "tanggal": tanggal,
                    "keterangan": keterangan,
                    "debet": debet,
                    "kredit": kredit,
                    "total_saldo": total_saldo
                }).execute()

                QMessageBox.information(self, "Sukses", "Data berhasil ditambahkan!")
                self.load_data()

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Gagal menambah data: {e}")

    # ====================================================
    # EDIT DATA (DENGAN FORM INPUT)
    # ====================================================
    def edit_data(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Peringatan", "Pilih data yang ingin diedit!")
            return

        try:
            id_data = self.table.item(row, 1).text()  # kolom 1 = id transaksi (lihat setHorizontalHeaderLabels)
            tanggal_lama = self.table.item(row, 2).text()
            keterangan_lama = self.table.item(row, 3).text()
            debet_lama = self.table.item(row, 4).text().replace(".", "").replace(",", ".")
            kredit_lama = self.table.item(row, 5).text().replace(".", "").replace(",", ".")

            # === Buat dialog input edit ===
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Data Transaksi")
            layout = QVBoxLayout(dialog)

            lbl_tanggal = QLabel("Tanggal Transaksi:")
            date_edit = QDateEdit(QDate.fromString(tanggal_lama, "yyyy-MM-dd"))
            date_edit.setCalendarPopup(True)

            lbl_keterangan = QLabel("Keterangan:")
            txt_keterangan = QLineEdit()
            txt_keterangan.setText(keterangan_lama)

            lbl_debet = QLabel("Uang Masuk (Rp):")
            txt_debet = QLineEdit()
            txt_debet.setText(str(debet_lama))

            lbl_kredit = QLabel("Uang Keluar (Rp):")
            txt_kredit = QLineEdit()
            txt_kredit.setText(str(kredit_lama))

            layout.addWidget(lbl_tanggal)
            layout.addWidget(date_edit)
            layout.addWidget(lbl_keterangan)
            layout.addWidget(txt_keterangan)
            layout.addWidget(lbl_debet)
            layout.addWidget(txt_debet)
            layout.addWidget(lbl_kredit)
            layout.addWidget(txt_kredit)

            btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            layout.addWidget(btn_box)
            btn_box.accepted.connect(dialog.accept)
            btn_box.rejected.connect(dialog.reject)

            # === Jalankan dialog ===
            if dialog.exec_() == QDialog.Accepted:
                tanggal_baru = date_edit.date().toString("yyyy-MM-dd")
                keterangan_baru = txt_keterangan.text().strip()
                debet_baru = txt_debet.text().strip()
                kredit_baru = txt_kredit.text().strip()

                # Validasi
                try:
                    debet_baru = float(debet_baru) if debet_baru else 0.0
                    kredit_baru = float(kredit_baru) if kredit_baru else 0.0
                except ValueError:
                    QMessageBox.warning(self, "Peringatan", "Masukkan angka yang valid untuk Uang Masuk / Keluar!")
                    return

                if not keterangan_baru:
                    QMessageBox.warning(self, "Peringatan", "Keterangan harus diisi!")
                    return

                # Ambil saldo sebelum transaksi ini
                res_before = supabase.table("pembukuan") \
                    .select("total_saldo") \
                    .lt("tanggal", tanggal_baru) \
                    .order("tanggal", desc=True) \
                    .limit(1) \
                    .execute().data
                saldo_sebelumnya = float(res_before[0]['total_saldo']) if res_before else 0

                # Hitung saldo baru untuk transaksi ini
                total_saldo_baru = saldo_sebelumnya + (debet_baru - kredit_baru)

                # Update ke Supabase
                supabase.table("pembukuan").update({
                    "tanggal": tanggal_baru,
                    "keterangan": keterangan_baru,
                    "debet": debet_baru,
                    "kredit": kredit_baru,
                    "total_saldo": total_saldo_baru
                }).eq("id", id_data).execute()

                QMessageBox.information(self, "Sukses", "Data berhasil diperbarui!")

                # === Perbarui saldo total setelah baris ini ===
                # Ambil semua data setelah tanggal transaksi ini, urut naik
                res_after = supabase.table("pembukuan") \
                    .select("id, tanggal, debet, kredit") \
                    .gt("tanggal", tanggal_baru) \
                    .order("tanggal", desc=False) \
                    .execute().data

                saldo_berjalan = total_saldo_baru
                for item in res_after:
                    saldo_berjalan += float(item["debet"]) - float(item["kredit"])
                    supabase.table("pembukuan").update({
                        "total_saldo": saldo_berjalan
                    }).eq("id", item["id"]).execute()

                self.load_data()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal mengedit data: {e}")

        # ====================================================
    # HAPUS DATA
    # ====================================================
    def hapus_data(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Peringatan", "Pilih data yang ingin dihapus!")
            return

        id_data = self.table.item(row, 1).text()  # kolom ID transaksi
        tanggal_data = self.table.item(row, 2).text()

        konfirmasi = QMessageBox.question(
            self,
            "Konfirmasi",
            f"Apakah Anda yakin ingin menghapus data tanggal {tanggal_data}?",
            QMessageBox.Yes | QMessageBox.No
        )
        if konfirmasi == QMessageBox.No:
            return

        try:
            # Hapus data dari Supabase
            supabase.table("pembukuan").delete().eq("id", id_data).execute()

            # Rehitung ulang total saldo setelah data ini
            # Ambil saldo terakhir sebelum tanggal ini
            res_before = supabase.table("pembukuan") \
                .select("total_saldo") \
                .lt("tanggal", tanggal_data) \
                .order("tanggal", desc=True) \
                .limit(1) \
                .execute().data
            saldo_sebelumnya = float(res_before[0]['total_saldo']) if res_before else 0

            # Ambil semua transaksi setelah tanggal ini
            res_after = supabase.table("pembukuan") \
                .select("id, tanggal, debet, kredit") \
                .gt("tanggal", tanggal_data) \
                .order("tanggal", desc=False) \
                .execute().data

            saldo_berjalan = saldo_sebelumnya
            for item in res_after:
                saldo_berjalan += float(item["debet"]) - float(item["kredit"])
                supabase.table("pembukuan").update({
                    "total_saldo": saldo_berjalan
                }).eq("id", item["id"]).execute()

            QMessageBox.information(self, "Sukses", "Data berhasil dihapus!")
            self.load_data()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal menghapus data: {e}")
    
    
    # ====================================================
    # LIHAT PERIODIK
    # ====================================================
    def lihat_periodik(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Lihat Data Periodik")
        layout = QVBoxLayout(dialog)

        lbl_awal = QLabel("Tanggal Awal:")
        lbl_akhir = QLabel("Tanggal Akhir:")
        date_awal = QDateEdit(QDate.currentDate().addMonths(-1))
        date_akhir = QDateEdit(QDate.currentDate())
        for d in (date_awal, date_akhir):
            d.setCalendarPopup(True)

        layout.addWidget(lbl_awal)
        layout.addWidget(date_awal)
        layout.addWidget(lbl_akhir)
        layout.addWidget(date_akhir)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)

        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)

        if dialog.exec_() == QDialog.Accepted:
            start_date = date_awal.date().toString("yyyy-MM-dd")
            end_date = date_akhir.date().toString("yyyy-MM-dd")

            res = supabase.table("pembukuan") \
                .select("total_saldo") \
                .lt("tanggal", start_date) \
                .order("tanggal", desc=True) \
                .limit(1) \
                .execute().data
            saldo_sebelumnya = float(res[0]['total_saldo']) if res else 0
            self.load_data(start_date=start_date, end_date=end_date, saldo_sebelumnya=saldo_sebelumnya)

    # ====================================================
    # LIHAT TAHUNAN (REKAP BULANAN + PILIHAN TAHUN + EKSPOR EXCEL)
    # ====================================================
    def lihat_tahunan(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Lihat Data Tahunan")
        layout = QVBoxLayout(dialog)

        lbl_tahun = QLabel("Pilih Tahun:")
        cmb_tahun = QComboBox()

        # Ambil semua tahun yang ada di tabel pembukuan (distinct)
        try:
            all_data = supabase.table("pembukuan").select("tanggal").order("tanggal", desc=False).execute().data
            # ekstrak tahun unik dari kolom tanggal
            tahun_unik = sorted({int(d["tanggal"][:4]) for d in all_data})
            cmb_tahun.clear()  # bersihkan dulu combo
            for th in tahun_unik:
                cmb_tahun.addItem(str(th))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Gagal memuat daftar tahun: {e}")
            return

        layout.addWidget(lbl_tahun)
        layout.addWidget(cmb_tahun)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)

        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)

        if dialog.exec_() == QDialog.Accepted:
            year = cmb_tahun.currentText()
            if not year:
                QMessageBox.warning(self, "Peringatan", "Tidak ada tahun yang dipilih!")
                return

            # ambil saldo sebelum tahun ini
            res = supabase.table("pembukuan") \
                .select("total_saldo") \
                .lt("tanggal", f"{year}-01-01") \
                .order("tanggal", desc=True) \
                .limit(1) \
                .execute().data
            saldo_awal = float(res[0]['total_saldo']) if res else 0

            # ambil semua data tahun ini (pakai range, bukan like)
            data_tahun = supabase.table("pembukuan") \
                .select("*") \
                .gte("tanggal", f"{year}-01-01") \
                .lte("tanggal", f"{year}-12-31") \
                .order("tanggal", desc=False) \
                .execute().data

            # hitung per bulan
            import calendar
            bulan_data = []
            saldo_akhir = saldo_awal
            for bulan in range(1, 13):
                nama_bulan = calendar.month_name[bulan]
                debet_bulan = sum(float(d["debet"]) for d in data_tahun if int(d["tanggal"][5:7]) == bulan)
                kredit_bulan = sum(float(d["kredit"]) for d in data_tahun if int(d["tanggal"][5:7]) == bulan)
                saldo_akhir = saldo_akhir + (debet_bulan - kredit_bulan)
                bulan_data.append({
                    "ID": bulan,
                    "Keterangan (Bulan)": nama_bulan,
                    "Saldo Debet": debet_bulan,
                    "Saldo Kredit": kredit_bulan,
                    "Saldo Akhir": saldo_akhir
                })

            # tampilkan di tabel
            self.table.setColumnCount(5)
            self.table.setHorizontalHeaderLabels(["ID", "Keterangan (Bulan)", "Saldo Debet", "Saldo Kredit", "Saldo Akhir"])
            self.table.setRowCount(len(bulan_data))

            for i, item in enumerate(bulan_data):
                self.table.setItem(i, 0, QTableWidgetItem(str(item["ID"])))
                self.table.setItem(i, 1, QTableWidgetItem(item["Keterangan (Bulan)"]))
                #self.table.setItem(i, 2, QTableWidgetItem(self.format_rupiah(item['Saldo Debet'])))
                #self.table.setItem(i, 3, QTableWidgetItem(self.format_rupiah(item['Saldo Kredit'])))
                #self.table.setItem(i, 4, QTableWidgetItem(self.format_rupiah(item['Saldo Akhir'])))
                item_debet = QTableWidgetItem(self.format_rupiah(item['Saldo Debet']))
                item_debet.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(i, 2, item_debet)

                item_kredit = QTableWidgetItem(self.format_rupiah(item['Saldo Kredit']))
                item_kredit.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(i, 3, item_kredit)

                item_saldo = QTableWidgetItem(self.format_rupiah(item['Saldo Akhir']))
                item_saldo.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(i, 4, item_saldo)

            self.lbl_title.setText(f"📆 Rekap Tahunan {year}")
            self.lbl_saldo_sebelumnya.setText(f"💼 Saldo Awal Tahun {year} = Rp {saldo_awal:,.2f}")
            self.lbl_ringkasan.setText(f"Saldo Akhir Tahun {year} = Rp {saldo_akhir:,.2f}")

            # ====== Tombol ekspor ke Excel ======
        if not hasattr(self, 'btn_ekspor'):
            self.btn_ekspor = QPushButton("⬇️ Ekspor ke Excel")
            self.btn_ekspor.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50; 
                    color: white; 
                    border: none; 
                    padding: 8px 14px;
                    border-radius: 5px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #43A047;
                }
            """)
            self.layout().addWidget(self.btn_ekspor)

            def ekspor_excel():
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "Simpan Laporan", f"Rekap_Tahunan_{year}.xlsx", "Excel Files (*.xlsx)"
                )
                if file_path:
                    df = pd.DataFrame(bulan_data)
                    df.to_excel(file_path, index=False, engine="openpyxl")
                    QMessageBox.information(self, "Sukses", f"Laporan berhasil diekspor ke:\n{file_path}")

            self.btn_ekspor.clicked.connect(ekspor_excel)


# ====================================================
# MAIN PROGRAM
# ====================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PembukuanApp()
    window.show()
    sys.exit(app.exec_())
