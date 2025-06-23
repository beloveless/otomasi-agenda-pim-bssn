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

# Mendapatkan tanggal hari ini
today = datetime.date.today()
today_sheet=today.strftime('%d').lstrip('0')

worksheet_name = f'24'  # Ganti dengan nama worksheet Anda
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"Worksheet dengan nama '{worksheet_name}' tidak ditemukan.")
    exit(1)

rows_to_clear = [6, 7, 13, 14, 20, 21, 27, 28]
columns_to_clear = range(2, 9)  # Kolom 2 hingga kolom 8

# Loop melalui setiap kolom yang ingin dihapus
for col in columns_to_clear:
    # Loop melalui setiap baris yang ingin dihapus
    for row in rows_to_clear:
    # Mengosongkan sel
        worksheet.update_cell(row, col, "")