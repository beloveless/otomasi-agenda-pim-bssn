import gspread,datetime,requests,json
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from datetime import datetime as dt
import os

# Tentukan scope yang dibutuhkan
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Masukkan path ke file JSON kredensial Anda
creds_path = os.path.join(os.path.dirname(__file__), 'teamup-425709-8333962f3c28.json')
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)

# Autentikasi dan inisialisasi client gspread
client = gspread.authorize(creds)

# Buka spreadsheet menggunakan ID
spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
spreadsheet = client.open_by_key(spreadsheet_id)

# Mendapatkan tanggal hari ini & besok
today = datetime.date.today()
tomorow = today + datetime.timedelta(days=1)
day_index = tomorow.weekday()

# Mengubah tanggal ke format yang diinginkan (YYYY-MM-DD)
tomorow_str=tomorow.strftime('%Y-%m-%d')
tomorow_sheet = tomorow.strftime("%d").lstrip('0')

worksheet_name = f'{tomorow_sheet}'  # Ganti dengan nama worksheet Anda
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"Worksheet dengan nama '{worksheet_name}' tidak ditemukan.")
    exit(1)

# Mendefinisikan range dari B7 ke H7
cell_range = 'B7:J7'

# Mengambil nilai dari sel-sel di range yang ditentukan
cells = worksheet.batch_get([cell_range])

# Mendefinisikan range dari B7 ke H7
start_col = 2  # Kolom B
end_col = 8    # Kolom H
row_index = 7  # Baris 7

# Memeriksa setiap sel dari B7 ke H7
for col_index in range(start_col, end_col + 1):
    cell_value = worksheet.cell(row_index, col_index).value
    if col_index ==2 or col_index==5 or col_index==6:
        kantor='Ragunan'
    else:
        kantor='Sawangan'

    if cell_value is None:
        if day_index==4:
            worksheet.update_cell(6, col_index, "08.00-16.30")
        else:
            worksheet.update_cell(6, col_index, "08.00-16.00")
        worksheet.update_cell(7, col_index, f"Berdinas Di Kantor {kantor}")