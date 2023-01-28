import streamlit as st
import pandas as pd
import datetime
from datetime import date
from pathlib import Path
import zipfile
import warnings
warnings.filterwarnings('ignore')
from openpyxl import load_workbook
from openpyxl.styles.alignment import Alignment

from functions import *


st.set_page_config(page_title="Mini Payroll System", layout="wide")
with st.sidebar:
#    st.sidebar.title(":abacus: Mini Payroll System")

    attendance_data = st.sidebar.file_uploader("**Upload Data Absensi**", type=["xlsx", "xls"])

    start_date = st.date_input(
        "**Start Date**",
        date.today())
    end_date = st.date_input(
        "**End Date**",
        date.today())
    st.markdown("""---""") 
    employee_master = st.sidebar.file_uploader("**Upload Master Data Pegawai**", type=["xlsx", "xls"])
    holidays_date = st.sidebar.file_uploader("**Upload Data Libur & Cuti Bersama**", type=["xlsx", "xls"])   
    denda_scan_masuk = st.number_input('**Denda Tidak Scan Masuk**',value=25000)
    denda_scan_pulang = st.number_input('**Denda Tidak Scan Pulang**',value=25000)
    uang_makan = st.number_input('**Uang Makan Harian**',value=15000)


font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
  font-size: 20px;
  font-weight: bold;
  margin-top: 0%;
}
</style>
"""

st.write(font_css, unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["Raw Data ", "Hitung Gaji Harian", "Generate Report Jam Kerja"])
with tab1:   

   st.write("#### Data Absensi")
   if attendance_data is not None:
        attendance_data_df = pd.read_excel(attendance_data)
        st.write(attendance_data_df)
   else:st.warning('Data Absensi belum di-upload', icon="⚠️")

   emp_master_last_updated = Path('emp_master_last_updated.txt').read_text()
   st.write("#### Data Master Pegawai")
   st.write("last update:",emp_master_last_updated)
   if employee_master is not None:
        employee_master_df = pd.read_excel(employee_master, dtype={'Nomor Rekening': str})
        employee_master_df.to_csv("temp_employee_master.csv",index=None)
        f = open('emp_master_last_updated.txt','w')
        f.write(str(datetime.now().strftime("%Y-%m-%d %H:%M")))
        f.close()
        st.write(employee_master_df)
   else:
        my_file = Path("temp_employee_master.csv")
        if my_file.is_file():
            st.write(pd.read_csv(my_file, dtype={'Nomor Rekening': str}))
        else:st.warning('Data Master Pegawai belum di-upload', icon="⚠️")

   holidays_date_last_updated = Path('holidays_date_last_updated.txt').read_text()
   st.write("#### Data Libur & Cuti Bersama")
   st.write("last update:",holidays_date_last_updated)
   if holidays_date is not None:
        holidays_date_df = pd.read_excel(holidays_date)
        holidays_date_df.to_csv("temp_holidays_date.csv",index=None)
        f = open('holidays_date_last_updated.txt','w')
        f.write(str(datetime.now().strftime("%Y-%m-%d %H:%M")))
        f.close()
        st.write(holidays_date_df)
   else:
        my_file = Path("temp_holidays_date.csv")
        if my_file.is_file():
            st.write(pd.read_csv(my_file))
        else:st.warning('Data Libur & Cuti Bersama belum di-upload', icon="⚠️")


with tab2:
   
   employee_master_df = pd.read_csv("temp_employee_master.csv", dtype={'Nomor Rekening': str})  
   holidays_date_df = pd.read_csv("temp_holidays_date.csv") 
   st.markdown(" \n\n")
   #st.write("Klik tombol dibawah untuk hitung gaji & generate kwitansi")

   col1, col2 = st.columns(2)

   with col1:
    btnHitungGaji = st.button("Hitung Gaji",help="klik tombol ini untuk hitung gaji",type="primary")

   if btnHitungGaji:
    with st.spinner('Loading....'):
        attendance_data_df["Tanggal"] = pd.to_datetime(attendance_data_df['Tanggal'])
        attendance_data_df["Tanggal"] = attendance_data_df["Tanggal"].dt.date
        filtered_attendance_df = attendance_data_df[attendance_data_df['Tanggal'].between(start_date, end_date)]
        # join to get payroll details
        absensi_emp_master = filtered_attendance_df.merge(employee_master_df[["PIN/ID","Keterangan","Gaji Harian (Pokok)","Upah Lembur","Nama Bank","Nama Akun Bank","Nomor Rekening"]],left_on='NIP', right_on='PIN/ID', how='left')
        absensi_emp_master['Tanggal'] = pd.to_datetime(absensi_emp_master['Tanggal'])
        holidays_date_df['Tanggal Libur'] = pd.to_datetime(holidays_date_df['Tanggal Libur'])
        # get holiday flag
        absensi_emp_master["weekday"] = absensi_emp_master["Tanggal"].dt.day_name()

        absensi_emp_master = absensi_emp_master.merge(holidays_date_df, left_on='Tanggal', right_on='Tanggal Libur',how='left')
        absensi_emp_master['is_holiday'] = absensi_emp_master['Tanggal'].isin(holidays_date_df['Tanggal Libur']).apply(lambda x: 'Y' if x else 'N')
        absensi_emp_master = absensi_emp_master.drop(columns=["PIN/ID","Tanggal Libur"])
        #st.write(absensi_emp_master)
        # get daily worker only data
        pekerja_harian = absensi_emp_master[absensi_emp_master["Keterangan"]=="HARIAN"]
        pekerja_harian_scan = pekerja_harian.drop(columns=["Tidak Scan Masuk","Tidak Scan Pulang"])

        # convert Scan related column into datetime
        for col in list(pekerja_harian_scan.filter(regex="Scan ").columns):
            pekerja_harian[col] = pd.to_datetime(pekerja_harian_scan[col], format="%H:%M:%S")

        pekerja_harian["Tanggal"] = pd.to_datetime(pekerja_harian["Tanggal"])

        # get scan_masuk and scan_pulang 
        pekerja_harian['scan_masuk'] = pekerja_harian[list(pekerja_harian_scan.filter(regex="Scan").columns)].min(axis=1)
        pekerja_harian['scan_pulang'] = pekerja_harian[list(pekerja_harian_scan.filter(regex="Scan").columns)].max(axis=1)

        #daily worker early scan (before 8AM)
        pekerja_harian['scan_masuk'] = pekerja_harian['scan_masuk'].apply(lambda x: pd.Timestamp('1900-01-01T08') if x<= pd.Timestamp('1900-01-01T08') else x)
        
        # get denda and uang makan 
        pekerja_harian["denda_tidak_scan_masuk"] = pekerja_harian["Tidak Scan Masuk"].apply(lambda x: denda_scan_masuk if x == "Y" else 0)
        pekerja_harian["denda_tidak_scan_pulang"] = pekerja_harian["Tidak Scan Pulang"].apply(lambda x: denda_scan_masuk if x == "Y" else 0)
        pekerja_harian["uang_makan_harian"] = pekerja_harian["Uang Makan"].apply(lambda x: uang_makan if x == "Y" else 0)

        #calculate working hours
        pekerja_harian[["jam_kerja", "jam_lembur","timedelta"]] = pekerja_harian.apply(calculate_work_hours, axis=1, result_type="expand")
        
        pekerja_harian[["gaji_harian", "gaji_lembur","total_gaji_harian"]] = pekerja_harian.apply(calculate_salary, axis=1, result_type="expand")

        total_gaji_df = pekerja_harian.groupby('NIP').agg({'total_gaji_harian': 'sum'}).rename(columns={'total_gaji_harian':'gaji_final'}).reset_index()

        gaji_pekerja_harian_details = pd.merge(pekerja_harian, total_gaji_df, on='NIP', how='left')

        df_kwitansi = gaji_pekerja_harian_details[["NIP","Nama","Nama Bank","Nama Akun Bank","Nomor Rekening","gaji_final"]].drop_duplicates().reset_index(drop=True)
        df_kwitansi["start_date"] = start_date
        df_kwitansi["end_date"] = end_date
        df_kwitansi[['nama_worksheet']] = df_kwitansi[['Nama']].replace(' ', '_', regex=True)

        file_list = generate_kwitansi(df_kwitansi)
        
        st.markdown("### Detail Gaji Pekerja Harian (preview)")
        st.write(gaji_pekerja_harian_details, formatter="{:.2f}")
        st.markdown("### Detail Kwitansi")
        st.write(df_kwitansi, formatter="{:.2f}")
        gaji_pekerja_harian_details.to_excel("kwitansi_output/"+"detail_perhitungan_gaji.xlsx",index=None)
        df_kwitansi.to_excel("kwitansi_output/"+"detail_kwitansi.xlsx",index=None)

        file_list.append("kwitansi_output/"+"detail_kwitansi.xlsx")
        file_list.append("kwitansi_output/"+"detail_perhitungan_gaji.xlsx")

        with zipfile.ZipFile("kwitansi_output/"+'kwitansi_'+str(date.today().strftime('%d%b%Y'))+'.zip', 'w') as zipMe:        
            for file in file_list:
                zipMe.write(file, compress_type=zipfile.ZIP_DEFLATED)
   with col2:
    my_file = Path("kwitansi_output/"+'kwitansi_'+str(date.today().strftime('%d%b%Y'))+'.zip')
    if my_file.is_file():
        with open("kwitansi_output/"+'kwitansi_'+str(date.today().strftime('%d%b%Y'))+'.zip', "rb") as fp:
            btn = st.download_button(
                label="Download Kwitansi",
                data=fp,
                file_name='kwitansi_'+str(date.today().strftime('%d%b%Y'))+'.zip',
                mime="application/zip"
            )


with tab3:
   st.header("TBU")
   st.write(employee_master_df)

