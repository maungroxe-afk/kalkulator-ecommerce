import streamlit as st
import pandas as pd
import io

# Pengaturan tampilan halaman
st.set_page_config(page_title="Kalkulator Harga E-Commerce Pro", layout="wide")

st.title("🛒 Kalkulator Harga Jual E-Commerce (Fitur Excel)")
st.write("Hitung harga jual ideal secara satuan atau masal (Bulk) agar terhindar dari kerugian biaya admin.")

# =====================================================================
# AREA UPDATE DATA BIAYA ADMIN & KOMISI (Update 2026)
# =====================================================================
FEE_DATA = {
    "Shopee": {
        "Fashion & Pakaian": {"fee": 0.10, "flat": 1250},
        "Kecantikan & Perawatan (termasuk Parfum)": {"fee": 0.0675, "flat": 1250},
        "Otomotif & Aksesoris Kendaraan": {"fee": 0.09, "flat": 1250},
        "Elektronik & Gadget Umum": {"fee": 0.095, "flat": 1250},
        "Elektronik High-End (HP, Laptop)": {"fee": 0.0525, "flat": 1250},
        "Makanan & Minuman (FMCG)": {"fee": 0.065, "flat": 1250}
    },
    "Tokopedia & TikTok Shop": {
        "Fashion (Pakaian, Sepatu, Tas)": {"fee": 0.08, "flat": 0},
        "Kecantikan & Perawatan Pribadi": {"fee": 0.07, "flat": 0},
        "Otomotif & Motor": {"fee": 0.075, "flat": 0},
        "Makanan & Minuman": {"fee": 0.065, "flat": 0},
        "Peralatan Rumah & Dapur": {"fee": 0.08, "flat": 0},
        "HP & Elektronik": {"fee": 0.03, "flat": 0}
    }
}

# Fungsi inti kalkulasi harga jual agar konsisten di satuan maupun bulk
def hitung_harga_jual_item(hpp, target_untung, biaya_lain, platform, kategori, is_po=False):
    if platform not in FEE_DATA or kategori not in FEE_DATA[platform]:
        return 0, 0
    
    persen_admin = FEE_DATA[platform][kategori]["fee"]
    biaya_tetap = FEE_DATA[platform][kategori]["flat"]
    
    if platform == "Shopee" and is_po:
        persen_admin += 0.03 # Tambahan komisi pre-order Shopee
        
    total_beban_dasar = hpp + target_untung + biaya_lain + biaya_tetap
    
    if persen_admin >= 1:
        return 0, 0
    
    # Rumus utama anti-rugi
    harga_jual = total_beban_dasar / (1 - persen_admin)
    potongan_persen = harga_jual * persen_admin
    
    # Aturan batas komisi maksimal Rp 650.000 (Tokopedia & TikTok Shop)
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
# MODE 2: UPLOAD MASAL VIA EXCEL
# ---------------------------------------------------------------------
elif menu == "Upload Masal (Bulk Excel)":
    st.header("📂 Perhitungan Banyak Produk Sekaligus (Bulk Upload)")
    st.write("Ikuti langkah di bawah ini untuk menghitung puluhan produk sekaligus menggunakan file Excel.")
    
    # LANGKAH 1: DOWNLOAD TEMPLATE
    st.subheader("Langkah 1: Download Template")
    st.write("Gunakan template standar di bawah ini agar sistem tidak error membaca data Anda.")
    
    # Data contoh untuk isi template
    template_df = pd.DataFrame({
        "Nama Produk": ["Kemeja Flanel X", "Parfum Scent Y"],
        "HPP": [85000, 45000],
        "Target Untung Bersih": [30000, 25000],
        "Biaya Packing & Lainnya": [2000, 1500],
        "Platform": ["Shopee", "Tokopedia & TikTok Shop"],
        "Kategori": ["Fashion & Pakaian", "Kecantikan & Perawatan (termasuk Parfum)"],
        "PreOrder Shopee (Ya/Tidak)": ["Tidak", "Tidak"]
    })
    
    # Tampilkan panduan penulisan teks kategori yang valid agar user tidak salah ketik
    with st.expander("Lihat Daftar Nama Platform & Kategori yang Valid"):
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.markdown("**Opsi untuk Platform 'Shopee':**")
            for k in FEE_DATA["Shopee"].keys():
                st.code(k)
        with col_info2:
            st.markdown("**Opsi untuk Platform 'Tokopedia & TikTok Shop':**")
            for k in FEE_DATA["Tokopedia & TikTok Shop"].keys():
                st.code(k)

    # Proses konversi dataframe contoh ke Excel di memori
    buffer_template = io.BytesIO()
    with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False, sheet_name="Template_Kalkulator")
    
    st.download_button(
        label="📥 Download Template Excel Kosong",
        data=buffer_template.getvalue(),
        file_name="template_kalkulator_masal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    
    # LANGKAH 2: UPLOAD & PROSES DATA
    st.subheader("Langkah 2: Upload File yang Sudah Diisi")
    uploaded_file = st.file_uploader("Pilih file Excel (.xlsx) Anda:", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            # Membaca excel yang diupload
            df = pd.read_excel(uploaded_file)
            
            # Cek kecocokan kolom utama
            required_cols = ["Nama Produk", "HPP", "Target Untung Bersih", "Biaya Packing & Lainnya", "Platform", "Kategori", "PreOrder Shopee (Ya/Tidak)"]
            if not all(col in df.columns for col in required_cols):
                st.error("Gagal membaca file! Pastikan nama kolom di file Anda sama persis dengan yang ada di template.")
            else:
                list_harga_jual = []
                list_potongan = []
                
                # Looping baris demi baris untuk dihitung harganya
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
                
                # Menyisipkan hasil kalkulasi ke kolom baru di data asli
                df["Harga Jual Disarankan (Rp)"] = list_harga_jual
                df["Potongan Admin Platform (Rp)"] = list_potongan
                df["Omset Kotor Per Produk (Rp)"] = df["Harga Jual Disarankan (Rp)"]
                
                st.success("🎉 Sukses! Seluruh data produk Anda berhasil dihitung.")
                st.dataframe(df, use_container_width=True)
                
                # LANGKAH 3: DOWNLOAD HASIL JADI
                st.subheader("Langkah 3: Simpan Hasil Perhitungan")
                
                buffer_result = io.BytesIO()
                with pd.ExcelWriter(buffer_result, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Hasil_Kalkulasi")
                
                st.download_button(
                    label="📥 Download Hasil Perhitungan (Excel)",
                    data=buffer_result.getvalue(),
                    file_name="hasil_kalkulasi_harga_jual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Terjadi error teknis saat membaca file Anda: {e}")
