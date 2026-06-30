import streamlit as st
import pandas as pd
import io
from openpyxl.worksheet.datavalidation import DataValidation

# Pengaturan tampilan halaman
st.set_page_config(page_title="Kalkulator Harga E-Commerce Pro", layout="wide")

st.title("🛒 Kalkulator Harga Jual E-Commerce (Dropdown Excel)")
st.write("Hitung harga jual ideal secara satuan atau masal menggunakan template Excel yang sudah dilengkapi fitur Dropdown Otomatis.")

# =====================================================================
# AREA UPDATE DATA BIAYA ADMIN & KOMISI (Update Terbaru 2026)
# Kategori disamakan fungsinya agar dropdown Excel sinkron & rapi
# =====================================================================
FEE_DATA = {
    "Shopee": {
        "Fashion & Pakaian": {"fee": 0.10, "flat": 1250},
        "Kecantikan & Perawatan (Parfum)": {"fee": 0.0675, "flat": 1250},
        "Otomotif & Aksesoris": {"fee": 0.09, "flat": 1250},
        "Elektronik & Gadget": {"fee": 0.095, "flat": 1250},
        "Makanan & Minuman": {"fee": 0.065, "flat": 1250},
        "Peralatan Rumah Tangga": {"fee": 0.06, "flat": 1250}
    },
    "Tokopedia & TikTok Shop": {
        "Fashion & Pakaian": {"fee": 0.08, "flat": 0},
        "Kecantikan & Perawatan (Parfum)": {"fee": 0.07, "flat": 0},
        "Otomotif & Aksesoris": {"fee": 0.075, "flat": 0},
        "Elektronik & Gadget": {"fee": 0.03, "flat": 0},
        "Makanan & Minuman": {"fee": 0.065, "flat": 0},
        "Peralatan Rumah Tangga": {"fee": 0.08, "flat": 0}
    }
}

# Fungsi inti kalkulasi harga jual
def hitung_harga_jual_item(hpp, target_untung, biaya_lain, platform, kategori, is_po=False):
    if platform not in FEE_DATA or kategori not in FEE_DATA[platform]:
        return 0, 0
    
    persen_admin = FEE_DATA[platform][kategori]["fee"]
    biaya_tetap = FEE_DATA[platform][kategori]["flat"]
    
    if platform == "Shopee" and is_po:
        persen_admin += 0.03 # Tambahan komisi pre-order Shopee 2026
        
    total_beban_dasar = hpp + target_untung + biaya_lain + biaya_tetap
    
    if persen_admin >= 1:
        return 0, 0
    
    harga_jual = total_beban_dasar / (1 - persen_admin)
    potongan_persen = harga_jual * persen_admin
    
    # Aturan batas komisi maksimal Rp 650.000 (Tokopedia & TikTok Shop 2026)
    if platform == "Tokopedia & TikTok Shop" and potongan_persen > 650000:
        potongan_persen = 650000
        harga_jual = hpp + target_untung + biaya_lain + biaya_tetap + potongan_persen
        
    total_potongan = potongan_persen + biaya_tetap
    return int(harga_jual), int(total_potongan)


# =====================================================================
# MENU NAVIGASI (SIDEBAR)
# =====================================================================
menu = st.sidebar.radio("Pilih Mode Perhitungan:", ["Kalkulator Satuan", "Upload Masal (Bulk Excel)"])

# ---------------------------------------------------------------------
# MODE 1: KALKULATOR SATUAN
# ---------------------------------------------------------------------
if menu == "Kalkulator Satuan":
    st.header("🧮 Kalkulator Satuan")
    
    col_hpp, col_untung = st.columns(2)
    with col_hpp:
        hpp = st.number_input("HPP (Modal Barang) (Rp)", min_value=0, value=150000, step=5000)
    with col_untung:
        target_untung = st.number_input("Target Untung Bersih (Rp)", min_value=0, value=35000, step=5000)

    biaya_lain = st.number_input("Biaya Ekstra (Packing, Lakban, dll) (Rp)", min_value=0, value=2500, step=500)

    col_plat, col_kat = st.columns(2)
    with col_plat:
        platform = st.selectbox("Marketplace:", list(FEE_DATA.keys()))
    with col_kat:
        kategori = st.selectbox("Kategori Produk:", list(FEE_DATA[platform].keys()))

    is_po = False
    if platform == "Shopee":
        is_po = st.checkbox("Barang ini menggunakan sistem Pre-Order (PO)?")

    if st.button("Hitung Harga Jual", type="primary", use_container_width=True):
        harga_jual, potongan = hitung_harga_jual_item(hpp, target_untung, biaya_lain, platform, kategori, is_po)
        
        st.divider()
        st.subheader("📊 Hasil Kalkulasi")
        st.metric("Tentukan Harga Jual Produk di Angka:", f"Rp {harga_jual:,}")
        
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"✅ Target Untung Anda: **Rp {target_untung:,}**")
            st.write(f"📦 Total Beban Modal (HPP + Packing): **Rp {hpp + biaya_lain:,}**")
        with c2:
            st.write(f"💸 Estimasi Potongan Marketplace: **Rp {potongan:,}**")

# ---------------------------------------------------------------------
# MODE 2: UPLOAD MASAL VIA EXCEL (DENGAN DROPDOWN OTOMATIS)
# ---------------------------------------------------------------------
elif menu == "Upload Masal (Bulk Excel)":
    st.header("📂 Perhitungan Banyak Produk Sekaligus (Bulk Upload)")
    st.write("Unduh template di bawah, isi menggunakan menu pilihan (dropdown) yang tersedia di Excel, lalu upload kembali.")
    
    st.subheader("Langkah 1: Download Template Ber-Dropdown")
    
    # 1. Membuat DataFrame Kosong/Contoh untuk Template
    template_df = pd.DataFrame({
        "Nama Produk": ["Produk Contoh A", "Produk Contoh B"],
        "HPP": [100000, 75000],
        "Target Untung Bersih": [25000, 20000],
        "Biaya Packing & Lainnya": [2000, 2000],
        "Platform": ["Shopee", "Tokopedia & TikTok Shop"],
        "Kategori": ["Fashion & Pakaian", "Kecantikan & Perawatan (Parfum)"],
        "PreOrder Shopee (Ya/Tidak)": ["Tidak", "Tidak"]
    })
    
    # 2. Proses pembuatan file Excel + Menyisipkan Fitur Dropdown
    buffer_template = io.BytesIO()
    with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False, sheet_name="Template_Kalkulator")
        
        # Ambil workbook dan worksheet openpyxl yang sedang aktif
        workbook = writer.book
        worksheet = writer.sheets["Template_Kalkulator"]
        
        # A. Buat Dropdown untuk PLATFORM (Kolom E)
        pilihan_platform = f'"{",".join(list(FEE_DATA.keys()))}"'
        dv_platform = DataValidation(type="list", formula1=pilihan_platform, allow_blank=True)
        dv_platform.error ='Pilihan tidak ada dalam daftar!'
        dv_platform.errorTitle = 'Pilihan Salah'
        dv_platform.prompt = 'Silakan pilih Platform'
        dv_platform.promptTitle = 'Platform'
        worksheet.add_data_validation(dv_platform)
        dv_platform.add("E2:E200") # Berlaku dari baris 2 sampai 200
        
        # B. Buat Dropdown untuk KATEGORI (Kolom F)
        # Mengambil daftar kategori unik dari struktur data
        daftar_kategori = list(FEE_DATA["Shopee"].keys())
        pilihan_kategori = f'"{",".join(daftar_kategori)}"'
        dv_kategori = DataValidation(type="list", formula1=pilihan_kategori, allow_blank=True)
        dv_kategori.error ='Kategori tidak ditemukan!'
        dv_kategori.errorTitle = 'Kategori Salah'
        dv_kategori.prompt = 'Silakan pilih Kategori Produk'
        dv_kategori.promptTitle = 'Kategori'
        worksheet.add_data_validation(dv_kategori)
        dv_kategori.add("F2:F200") # Berlaku dari baris 2 sampai 200
        
        # C. Buat Dropdown untuk PREORDER (Kolom G)
        dv_po = DataValidation(type="list", formula1='"Ya,Tidak"', allow_blank=True)
        dv_po.prompt = 'Apakah produk ini Pre-Order?'
        dv_po.promptTitle = 'Sistem PO'
        worksheet.add_data_validation(dv_po)
        dv_po.add("G2:G200")
        
        # Atur lebar kolom agar rapi saat dibuka di Excel
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    st.download_button(
        label="📥 Download Template Excel (Ada Dropdown)",
        data=buffer_template.getvalue(),
        file_name="template_kalkulator_dropdown.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    
    # LANGKAH 2: UPLOAD & PROSES DATA
    st.subheader("Langkah 2: Upload File yang Sudah Diisi")
    uploaded_file = st.file_uploader("Pilih file Excel yang sudah diisi:", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Validasi struktur kolom
            required_cols = ["Nama Produk", "HPP", "Target Untung Bersih", "Biaya Packing & Lainnya", "Platform", "Kategori", "PreOrder Shopee (Ya/Tidak)"]
            if not all(col in df.columns for col in required_cols):
                st.error("Struktur file salah! Gunakan kolom asli dari template yang diunduh di atas.")
            else:
                list_harga_jual = []
                list_potongan = []
                
                for idx, row in df.iterrows():
                    is_po_bool = True if str(row["PreOrder Shopee (Ya/Tidak)"]).strip().lower() == "ya" else False
                    
                    harga_jual, potongan = hitung_harga_jual_item(
                        hpp=int(row["HPP"]),
                        target_untung=int(row["Target Untung Bersih"]),
                        biaya_lain=int(row["Biaya Packing & Lainnya"]),
                        platform=str(row["Platform"]).strip(),
                        kategori=str(row["Kategori"]).strip(),
                        is_po=is_po_bool
                    )
                    list_harga_jual.append(harga_jual)
                    list_potongan.append(potongan)
                
                # Input data hasil kalkulasi ke dataframe
                df["Harga Jual Disarankan (Rp)"] = list_harga_jual
                df["Potongan Admin Platform (Rp)"] = list_potongan
                df["Estimasi Pendapatan Kotor (Rp)"] = df["Harga Jual Disarankan (Rp)"]
                
                st.success("🎉 Seluruh data berhasil dihitung otomatis tanpa typo!")
                st.dataframe(df, use_container_width=True)
                
                # LANGKAH 3: DOWNLOAD HASIL JADI
                st.subheader("Langkah 3: Simpan Hasil Perhitungan")
                
                buffer_result = io.BytesIO()
                with pd.ExcelWriter(buffer_result, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Hasil_Kalkulasi")
                
                st.download_button(
                    label="📥 Download Hasil Perhitungan (Excel)",
                    data=buffer_result.getvalue(),
                    file_name="hasil_kalkulasi_toko.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Terjadi kesalahan membaca file. Pastikan data angka terisi dengan benar. Error: {e}")
