import streamlit as st
import pandas as pd
import xlsxwriter
import io
from datetime import datetime  # Impor datetime

st.set_page_config(
    page_title="Data Jembatan Timbang",
)

# Membuat tata letak dengan st.columns untuk logo Kemenhub dan judul
col1, col2 = st.columns([1, 7])
with col1:
    st.image('logo_kemenhub.png', width=90)
with col2:
    st.title('Data Angkutan Barang Masuk Di Jembatan Timbang Cekik')

# Inisialisasi nama file Excel
excel_file = 'data_angkutan.xlsx'

# Fungsi untuk membaca data dari file Excel
def read_excel(file):
    try:
        df = pd.read_excel(file)
        df.index = df.index + 1  # Memulai indeks dari 1
    except FileNotFoundError:
        df = pd.DataFrame(columns=['No', 'Nomor Kendaraan', 'Nama Perusahaan', 'Jenis Muatan', 'Berat Muatan', 'Berat Kosong', 'JBI', 'Status', 'Tanggal'])
    return df

# Fungsi untuk menulis data ke file Excel
def write_excel(df, file):
    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')  # tanpa indeks pada sheet 1

# Membaca data dari file Excel saat memulai aplikasi
if 'transportation_data' not in st.session_state:
    st.session_state['transportation_data'] = read_excel(excel_file)

# Fungsi untuk menambah kendaraan
def add_vehicle(df, vehicle_number, driver, jenismuatan, beratmuatan, beratkosong, jbi):
    new_vehicle = pd.DataFrame({
        'No': [df.shape[0] + 1],  # Menggunakan jumlah baris + 1 untuk nomor urut
        'Nomor Kendaraan': [vehicle_number],
        'Nama Perusahaan': [driver],
        'Jenis Muatan': [jenismuatan],
        'Berat Muatan': [beratmuatan],
        'Berat Kosong': [beratkosong],
        'JBI': [jbi],
        'Status': [''],
        'Tanggal': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]  # Tambahkan kolom Tanggal dengan waktu saat ini
    })
    df = pd.concat([df, new_vehicle], ignore_index=True)
    df.index = df.index + 1  # Memulai indeks dari 1
    return df

# Fungsi untuk menghapus semua data
def clear_data():
    st.session_state['transportation_data'] = pd.DataFrame(columns=['No', 'Nomor Kendaraan', 'Nama Perusahaan', 'Jenis Muatan', 'Berat Muatan', 'Berat Kosong', 'JBI', 'Status', 'Tanggal'])
    write_excel(st.session_state['transportation_data'], excel_file)

# INPUT KENDARAAN BARU
st.subheader('Input Kendaraan Baru')
with st.form('Tambah Kendaraan'):
    vehicle_number = st.text_input('Nomor Kendaraan')
    driver = st.text_input('Nama Perusahaan')
    jenismuatan = st.text_input('Jenis Muatan')
    beratmuatan = st.text_input('Berat Muatan')
    beratkosong = st.text_input('Berat Kosong')
    jbi = st.text_input('JBI')
    submitted = st.form_submit_button('Tambah')

    if submitted:
        if vehicle_number in st.session_state['transportation_data']['Nomor Kendaraan'].values:
            st.error('Nomor kendaraan sudah ada!')
        else:
            st.session_state['transportation_data'] = add_vehicle(st.session_state['transportation_data'], vehicle_number, driver, jenismuatan, beratmuatan, beratkosong, jbi)
            write_excel(st.session_state['transportation_data'], excel_file)  # Menulis data ke file Excel
            st.success('Kendaraan berhasil ditambahkan!')

# UPDATE STATUS KENDARAAN
st.subheader('Status Kendaraan')
if not st.session_state['transportation_data'].empty:
    vehicle_to_update = st.selectbox('Pilih Nomor Kendaraan', st.session_state['transportation_data']['Nomor Kendaraan'].unique())
    
    # Periksa apakah data kendaraan valid sebelum melakukan update
    if st.button('Update Status Otomatis'):
        try:
            # Ambil data berat muatan dan JBI kendaraan yang dipilih
            berat_muatan = float(st.session_state['transportation_data'].loc[st.session_state['transportation_data']['Nomor Kendaraan'] == vehicle_to_update, 'Berat Muatan'].values[0])
            jbi = float(st.session_state['transportation_data'].loc[st.session_state['transportation_data']['Nomor Kendaraan'] == vehicle_to_update, 'JBI'].values[0])
            
            # Logika otomatis untuk memperbarui status
            if berat_muatan > jbi:
                new_status = "Over Load"
            else:
                new_status = "Aman"
            
            # Perbarui status di DataFrame
            st.session_state['transportation_data'].loc[st.session_state['transportation_data']['Nomor Kendaraan'] == vehicle_to_update, 'Status'] = new_status
            
            # Simpan kembali data ke file Excel
            write_excel(st.session_state['transportation_data'], excel_file)
            st.success(f'Status kendaraan berhasil diperbarui menjadi "{new_status}"!')
        
        except ValueError:
            st.error("Pastikan kolom Berat Muatan dan JBI diisi dengan angka yang valid.")
else:
    st.write("Tidak ada data kendaraan untuk diperbarui.")


# Tombol Clear Data
if st.button('Clear Data'):
    clear_data()
    st.success('Semua data telah dihapus!')

# EXPORT KE EXCEL
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='sheet1')

        # Akses workbook dan worksheet untuk formatting
        workbook = writer.book
        worksheet = writer.sheets['sheet1']

        # Format untuk header
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'center',
            'align': 'center',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # Terapkan format header
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Sesuaikan lebar kolom secara otomatis
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_width)

    processed_data = output.getvalue()
    return processed_data


if st.button('Download Data as Excel'):
    st.write('Excel File sudah bisa di Download')
    df_xlsx = to_excel(st.session_state['transportation_data'])
    st.download_button(label='Download Excel file',
                       data=df_xlsx,
                       file_name='data_angkutan.xlsx')

# Selalu tampilkan tabel
st.subheader('Data Kendaraan')
st.dataframe(st.session_state['transportation_data'])
