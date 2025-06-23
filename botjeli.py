import gspread,datetime,requests
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from datetime import datetime as dt
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
from telegram import Bot
import io

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds_path = os.path.join(os.path.dirname(__file__), 'teamup-425709-8333962f3c28.json')

# Autentikasi gspread
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)

# Autentikasi Google Drive API
drive_creds = service_account.Credentials.from_service_account_file(creds_path, scopes=scope)

# Buka spreadsheet menggunakan ID
# spreadsheet_id = '1aJmL60KFsqbTIRxgDoBJnVjiuOw5xTSMhr32-u7BWtw'
spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
spreadsheet = client.open_by_key(spreadsheet_id)

# Mendapatkan tanggal hari ini & besok
today = datetime.date.today()
tomorow = today + datetime.timedelta(days=1)

# Mengubah tanggal ke format yang diinginkan (YYYY-MM-DD)
tomorow_str=tomorow.strftime('%Y-%m-%d')
tomorow_sheet = tomorow.strftime("%d").lstrip('0')

worksheet_name = f'{tomorow_sheet}'  # Ganti dengan nama worksheet Anda
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"Worksheet dengan nama '{worksheet_name}' tidak ditemukan.")
    exit(1)

# Fungsi untuk memformat waktu dari string datetime
def format_time(start_datetime_str, end_datetime_str):
    # Mengonversi string datetime menjadi objek datetime
    start_dt = dt.fromisoformat(start_datetime_str)
    end_dt = dt.fromisoformat(end_datetime_str)
    # Mengambil bagian waktu dan memformatnya ke HH.MM
    start_time = start_dt.strftime("%H.%M")
    end_time = end_dt.strftime("%H.%M")
    return f"{start_time} - {end_time}"

# Fungsi untuk mengambil data dari URL Teamup
def get_teamup_data(url, token):

    headers = {
        'Teamup-Token': token
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # Memastikan permintaan berhasil
    if response.status_code == 200:
        return response.json()
    else:
        print("Gagal mengambil data dari Teamup.")
        return None
    
# Fungsi untuk menghapus kolom kosong dari worksheet
def remove_empty_columns(ws):
    data = ws.get_all_values()
    if not data:
        return
    num_cols = len(data[0])
    for col in range(num_cols, 0, -1):
        if all((row[col-1] == '' for row in data[5:])):
            ws.delete_columns(col)

# URL Teamup yang ingin Anda akses
teamup_url = f'https://api.teamup.com/ksjvi17ce1ipimpco8/events?startDate={tomorow_str}&endDate={tomorow_str}&tz=Jakarta' #reader pada tanggal tertentu
api_token = '7c35b104dc19267c81d4914d2e70159f1b9f5274a4e612e4c6e0a1fd766b54ab'

# Ambil data dari Teamup
teamup_data = get_teamup_data(teamup_url,api_token)

# Kamus untuk memetakan sub calendar ke kolom
subcalendar_to_col = {
    10858904: 2,
    10859020: 3,
    10860315: 4,
    10859016: 5,
    10859017: 6,
    10859018: 7,
    10859019: 8
}

# Kamus untuk melacak baris awal untuk setiap sub calendar (dimulai dari baris ke-3)
subcalendar_to_row = {sub_calendar: 6 for sub_calendar in range(1, 8)}


if teamup_data:

    # Ambil data yang diperlukan dari JSON
    events = teamup_data["events"]

    # Mulai dari baris ke-3 untuk setiap subkalender
    subcalendar_to_row = {sub_calendar: 6 for sub_calendar in subcalendar_to_col}

    for event in events:
        # Ambil data dari JSON
        sub_calendar = event["subcalendar_id"]
        start_date = event["start_dt"]
        end_date = event["end_dt"]
        event_title = event["title"]
        event_location = event.get("location", "")

        # Tentukan kolom berdasarkan sub calendar
        col = subcalendar_to_col.get(sub_calendar, None)
        if col is None:
            continue  # Jika subcalendar tidak ada dalam kamus, lewati event ini
        
        # Kumpulkan semua sub_calendar_id dari events
        event_subcalendars = {event["subcalendar_id"] for event in events}

         # Ambil baris awal untuk sub calendar ini
        start_row = subcalendar_to_row[sub_calendar]

        # Format start date ke dalam format waktu HH.MM
        formatted_time = format_time(start_date, end_date)

        # Update sel dengan start date dan enddate
        worksheet.update_cell(start_row, col, formatted_time)
        start_row += 1

         # Periksa apakah lokasi ada dan tidak kosong
        if event_location:
            event_title_with_location = f"{event_title} di {event_location}"
        else:
            event_title_with_location = event_title

        # Update sel dengan judul acara
        worksheet.update_cell(start_row, col, event_title_with_location)
        
         # Update baris awal untuk sub kalender ini
        subcalendar_to_row[sub_calendar] = start_row + 6

# Hapus kolom kosong setelah semua data dimasukkan
remove_empty_columns(worksheet)

# Ekspor sebagai PDF dan kirim ke Telegram
pdf_file_name = f"agenda_{tomorow_str}.pdf"
drive_service = build('drive', 'v3', credentials=drive_creds)
request = drive_service.files().export_media(
    fileId=spreadsheet_id, 
    mimeType='application/pdf')
request.uri += "&portrait=false"  # Set landscape mode


fh = io.BytesIO()
downloader = MediaIoBaseDownload(fh, request)
done = False
while done is False:
    status, done = downloader.next_chunk()

# Simpan file PDF
with open(pdf_file_name, 'wb') as f:
    f.write(fh.getvalue())

# Kirim PDF ke Telegram
bot_token = '7135848774:AAGlJLq4eR4WXaS6k8KHL4RoJkywJ09b7mo'
chat_id = '642574222'  # ID Telegram untuk @itsenjelly
bot = Bot(token=bot_token)
with open(pdf_file_name, 'rb') as file:
    bot.send_document(chat_id=chat_id, document=file, filename=pdf_file_name)

print("Data dimigrasikan, file PDF dikirim ke Telegram.")