import gspread, datetime, requests, os, json
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from datetime import datetime as dt
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from telegram import Bot
import asyncio

# === Debug: Periksa variabel environment ===
bot_token = os.getenv('TELEGRAM_TOKEN')
chat_id = os.getenv('CHAT_ID')
teamup_token = os.getenv('TEAMUP_TOKEN')
google_creds = os.getenv('GOOGLE_CREDS')

print("🔧 Cek variabel environment:")
print(f"- TELEGRAM_TOKEN: {'✅ Ada' if bot_token else '❌ Kosong'}")
print(f"- CHAT_ID: {chat_id}")
print(f"- TEAMUP_TOKEN: {'✅ Ada' if teamup_token else '❌ Kosong'}")
print(f"- GOOGLE_CREDS: {'✅ Ada' if google_creds else '❌ Kosong'}")

# === Buat file kredensial dari secret ===
with open('creds.json', 'w') as f:
    f.write(google_creds)

# === Autentikasi Google API ===
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
client = gspread.authorize(creds)
drive_creds = service_account.Credentials.from_service_account_file('creds.json', scopes=scope)

# === Tanggal dan Worksheet ===
today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
tomorrow_str = tomorrow.strftime('%Y-%m-%d')
day_index = tomorrow.weekday()
tomorrow_sheet = tomorrow.strftime("%d").lstrip('0')
worksheet_name = f'{tomorrow_sheet}'

spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
spreadsheet = client.open_by_key(spreadsheet_id)

try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"❌ Worksheet '{worksheet_name}' tidak ditemukan.")
    exit(1)


try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"❌ Worksheet '{worksheet_name}' tidak ditemukan.")
    exit(1)

# === Tambahan: Menuliskan Hari dan Tanggal setelah Judul ===
def tulis_hari_dan_tanggal(ws, tanggal: datetime.date):
    hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'][tanggal.weekday()]
    tanggal_str = tanggal.strftime('%d %B %Y')
    keterangan = f"{hari}, {tanggal_str}"
    
    all_values = ws.get_all_values()
    for i, row in enumerate(all_values, start=1):
        if any("Agenda Kegiatan Pimpinan BSSN" in cell for cell in row):
            ws.update_cell(i+1, 2, keterangan)  # Baris di bawahnya, kolom B
            print(f"🗓️ Ditambahkan keterangan tanggal: {keterangan} di baris {i+1}")
            break

tulis_hari_dan_tanggal(worksheet, tomorrow)

# === Fungsi Format dan Migrasi ===
def format_time(start_datetime_str, end_datetime_str):
    start_dt = dt.fromisoformat(start_datetime_str)
    end_dt = dt.fromisoformat(end_datetime_str)
    return f"{start_dt.strftime('%H.%M')} - {end_dt.strftime('%H.%M')}"

def get_teamup_data(url, token):
    headers = {'Teamup-Token': token}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def add_rows_with_border(ws, count):
    last_row = len(ws.get_all_values())
    for _ in range(count):
        ws.insert_rows([[''] * 8 for _ in range(7)], row=last_row+1)
        for i in range(last_row+1, last_row+8):
            set_border(ws, f"B{i}:I{i}", style='SOLID', color=Color(0,0,0))
        last_row += 7

def remerge_and_number_blocks(ws):
    data = ws.get_all_values()
    for i in range(5, len(data), 7):
        try: ws.unmerge_cells(f"A{i+1}:A{i+7}")
        except: pass
        ws.update_cell(i+1, 1, str((i-5)//7 + 1))
        ws.merge_cells(f"A{i+1}:A{i+7}")

def remove_empty_agenda_blocks(ws):
    data = ws.get_all_values()
    for start_row in reversed(range(6, len(data)+1, 7)):
        block = data[start_row-1:start_row+6]
        if all(all(cell.strip() == '' for cell in row[1:8]) for row in block):
            for _ in range(7):
                ws.delete_rows(start_row)

def isi_jika_kosong(ws):
    for col_index in range(2, 9):
        cell_value = ws.cell(7, col_index).value
        kantor = 'Ragunan' if col_index in [2, 5, 6] else 'Sawangan'
        if not cell_value:
            jam = "08.00-16.30" if day_index == 4 else "08.00-16.00"
            ws.update_cell(6, col_index, jam)
            ws.update_cell(7, col_index, f"Berdinas Di Kantor {kantor}")

# === Proses Migrasi ===
subcalendar_to_col = {10858904: 2, 10859020: 3, 10860315: 4, 10859016: 5, 10859017: 6, 10859018: 7, 10859019: 8}
subcalendar_to_row = {k: 6 for k in subcalendar_to_col}

teamup_url = f'https://api.teamup.com/ksjvi17ce1ipimpco8/events?startDate={tomorrow_str}&endDate={tomorrow_str}&tz=Asia/Jakarta'
data = get_teamup_data(teamup_url, teamup_token)

if data:
    events = data.get("events", [])
    print(f"📅 Jumlah event ditemukan: {len(events)}")
    needed_blocks = (len(events) + 6) // 7
    current_rows = len(worksheet.get_all_values())
    current_blocks = (current_rows - 5) // 7
    if needed_blocks > current_blocks:
        add_rows_with_border(worksheet, needed_blocks - current_blocks)

    for e in events:
        col = subcalendar_to_col.get(e['subcalendar_id'])
        if not col: continue
        row = subcalendar_to_row[e['subcalendar_id']]
        worksheet.update_cell(row, col, format_time(e['start_dt'], e['end_dt']))
        worksheet.update_cell(row+1, col, f"{e['title']} di {e.get('location', '')}".strip())
        subcalendar_to_row[e['subcalendar_id']] = row + 7

isi_jika_kosong(worksheet)
remove_empty_agenda_blocks(worksheet)
remerge_and_number_blocks(worksheet)

# === Ekspor PDF ===
if not drive_creds.valid:
    drive_creds.refresh(Request())

export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=pdf&portrait=false&gridlines=false&size=A4&fitw=true&gid={worksheet._properties['sheetId']}"
headers = {"Authorization": f"Bearer {drive_creds.token}"}
pdf_file_name = f"agenda_{tomorrow_str}.pdf"
response = requests.get(export_url, headers=headers)

with open(pdf_file_name, 'wb') as f:
    f.write(response.content)

print(f"✅ PDF berhasil dibuat: {pdf_file_name}")
print(f"📁 Cek file ada? {os.path.exists(pdf_file_name)}")

# === Kirim ke Telegram ===
async def send_pdf():
    try:
        bot = Bot(token=bot_token)
        with open(pdf_file_name, 'rb') as file:
            await bot.send_document(chat_id=chat_id, document=file, filename=pdf_file_name)
        print("✅ PDF berhasil dikirim ke Telegram.")
    except Exception as e:
        print(f"❌ Gagal mengirim ke Telegram: {e}")

# Jalankan fungsi async
asyncio.run(send_pdf())