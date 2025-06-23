import gspread, datetime, requests, os, json, io
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from datetime import datetime as dt
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from telegram import Bot

# === Konfigurasi ===
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds_path = os.path.join(os.path.dirname(__file__), 'teamup-425709-8333962f3c28.json')
spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
bot_token = '7135848774:AAGlJLq4eR4WXaS6k8KHL4RoJkywJ09b7mo'
chat_id = '642574222'
teamup_token = '7c35b104dc19267c81d4914d2e70159f1b9f5274a4e612e4c6e0a1fd766b54ab'

# === Autentikasi Google API ===
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)
drive_creds = service_account.Credentials.from_service_account_file(creds_path, scopes=scope)

# === Tanggal dan Worksheet ===
today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
tomorrow_str = tomorrow.strftime('%Y-%m-%d')
day_index = tomorrow.weekday()
tomorrow_sheet = tomorrow.strftime("%d").lstrip('0')
worksheet_name = f'{tomorrow_sheet}'

spreadsheet = client.open_by_key(spreadsheet_id)
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"Worksheet '{worksheet_name}' tidak ditemukan.")
    exit(1)

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

# === Tahapan Lanjutan ===
isi_jika_kosong(worksheet)
remove_empty_agenda_blocks(worksheet)
remerge_and_number_blocks(worksheet)

if not drive_creds.valid:
    drive_creds.refresh(Request())

export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=pdf&portrait=false&gridlines=false&size=A4&fitw=true&gid={worksheet._properties['sheetId']}"
headers = {"Authorization": f"Bearer {drive_creds.token}"}
pdf_file_name = f"agenda_{tomorrow_str}.pdf"
response = requests.get(export_url, headers=headers)
with open(pdf_file_name, 'wb') as f:
    f.write(response.content)

# === Kirim ke Telegram ===
bot = Bot(token=bot_token)
with open(pdf_file_name, 'rb') as file:
    bot.send_document(chat_id=chat_id, document=file, filename=pdf_file_name)

print("Selesai: data migrasi, pengisian, penghapusan blok kosong, dan pengiriman PDF ke Telegram.")