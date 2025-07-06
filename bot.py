import gspread, datetime, requests, os, json, pytz
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from datetime import datetime as dt
from google.oauth2 import service_account
from google.auth.transport.requests import Request
import asyncio
from telegram.ext import ApplicationBuilder

# === Ambil Secrets dari Environment ===
try:
    google_creds = os.environ.get("GOOGLE_CREDS")
    bot_token = os.environ.get("TELEGRAM_TOKEN")
    chat_id = os.environ.get("CHAT_ID")
    teamup_token = os.environ.get("TEAMUP_TOKEN")
    if not all([google_creds, bot_token, chat_id, teamup_token]):
        raise ValueError("Beberapa environment variable tidak tersedia.")
    print("✅ Secrets berhasil di-load.")
except Exception as e:
    print(f"❌ Gagal load secrets: {e}")
    exit(1)

# === Simpan kredensial Google sementara ===
try:
    creds_path = "creds.json"
    with open(creds_path, "w") as f:
        f.write(google_creds)
    print("✅ File kredensial Google disimpan.")
except Exception as e:
    print(f"❌ Gagal menyimpan file kredensial: {e}")
    exit(1)

# === Autentikasi Google API ===
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
try:
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    client = gspread.authorize(creds)
    drive_creds = service_account.Credentials.from_service_account_file(creds_path, scopes=scope)
    print("✅ Autentikasi Google API berhasil.")
except Exception as e:
    print(f"❌ Gagal autentikasi Google API: {e}")
    exit(1)

# === Tanggal & Worksheet ===
try:
    jakarta = pytz.timezone("Asia/Jakarta")
    now_jakarta = datetime.datetime.now(jakarta)
    today = now_jakarta.date()
    tomorrow = today + datetime.timedelta(days=1)
    tomorrow_str = tomorrow.strftime('%Y-%m-%d')
    day_index = tomorrow.weekday()
    sheet_name = tomorrow.strftime("%d").lstrip('0')

    print(f"🕓 Sekarang (WIB): {now_jakarta.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📅 Target agenda tanggal: {tomorrow_str}")

    spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    print(f"✅ Worksheet '{sheet_name}' ditemukan.")
    hapus_semua_agenda(worksheet)
except gspread.WorksheetNotFound:
    print(f"❌ Worksheet '{sheet_name}' tidak ditemukan.")
    exit(1)
except Exception as e:
    print(f"❌ Gagal membuka worksheet: {e}")
    exit(1)

def tulis_hari_dan_tanggal(ws, tanggal: datetime.date):
    hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'][tanggal.weekday()]
    tanggal_str = tanggal.strftime('%d %B %Y')
    keterangan = f"{hari}, {tanggal_str}"
    ws.update('A2', [[keterangan]])
    ws.merge_cells('A2:H2')
    format_cell_range(ws, 'A2:H2', CellFormat(
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    print(f"🗓️ Ditambahkan keterangan tanggal di baris 2 (A2:H2): {keterangan}")

# === Fungsi Utilitas ===
def format_time(start, end):
    start_dt = dt.fromisoformat(start)
    end_dt = dt.fromisoformat(end)
    return f"{start_dt.strftime('%H.%M')} - {end_dt.strftime('%H.%M')}"

def get_teamup_data(url, token):
    try:
        headers = {'Teamup-Token': token}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ Gagal ambil data Teamup: {e}")
        return None

def add_rows_with_border(ws, count):
    last_row = len(ws.get_all_values())
    for _ in range(count):
        ws.insert_rows([[''] * 8 for _ in range(7)], row=last_row+1)
        for i in range(last_row+1, last_row+8):
            set_border(ws, f"B{i}:I{i}", style='SOLID', color=Color(0, 0, 0))
        last_row += 7

def remerge_and_number_blocks(ws):
    data = ws.get_all_values()
    for i in range(5, len(data), 7):
        try:
            ws.unmerge_cells(f"A{i+1}:A{i+7}")
        except:
            pass
        ws.update_cell(i+1, 1, str((i-5)//7 + 1))
        ws.merge_cells(f"A{i+1}:A{i+7}")

def remove_empty_agenda_blocks(ws):
    data = ws.get_all_values()
    for start_row in reversed(range(13, len(data)+1, 7)):
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

def hapus_semua_agenda(ws):
    data = ws.get_all_values()
    start_row = 6  # Baris awal agenda, sesuaikan jika beda
    if len(data) > start_row:
        end_col = chr(64 + len(data[0]))  # Misal: H = kolom ke-8
        ws.batch_clear([f"A{start_row}:{end_col}{len(data)}"])
        print(f"🧹 Semua blok agenda mulai dari baris {start_row} berhasil dihapus.")
    else:
        print("ℹ️ Worksheet sudah kosong dari baris agenda.")

# === Ambil dan isi data dari Teamup ===
subcalendar_to_col = {
    10858904: 2, 10859020: 3, 10860315: 4,
    10859016: 5, 10859017: 6, 10859018: 7, 10859019: 8
}
subcalendar_to_row = {k: 6 for k in subcalendar_to_col}

teamup_url = f'https://api.teamup.com/ksjvi17ce1ipimpco8/events?startDate={tomorrow_str}&endDate={tomorrow_str}&tz=Asia/Jakarta'
teamup_data = get_teamup_data(teamup_url, teamup_token)

if teamup_data:
    print("✅ Data dari Teamup berhasil diambil.")
    events = teamup_data.get("events", [])
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
    print("✅ Data berhasil dimasukkan ke worksheet.")
else:
    print("⚠️ Tidak ada data dari Teamup.")

# === Formatting worksheet ===
try:
    tulis_hari_dan_tanggal(worksheet, tomorrow)
    isi_jika_kosong(worksheet)
    remove_empty_agenda_blocks(worksheet)
    remerge_and_number_blocks(worksheet)
    print("✅ Worksheet berhasil diformat.")
except Exception as e:
    print(f"❌ Gagal memformat worksheet: {e}")

# === Ekspor PDF ===
try:
    if not drive_creds.valid:
        drive_creds.refresh(Request())
    pdf_file_name = f"agenda_{tomorrow_str}.pdf"
    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=pdf&portrait=false&gridlines=false&size=A4&fitw=true&gid={worksheet._properties['sheetId']}"
    headers = {"Authorization": f"Bearer {drive_creds.token}"}
    response = requests.get(export_url, headers=headers)
    with open(pdf_file_name, 'wb') as f:
        f.write(response.content)
    print("✅ PDF berhasil diekspor.")
except Exception as e:
    print(f"❌ Gagal ekspor PDF: {e}")
    exit(1)

# === Kirim ke Telegram ===
async def send_telegram_message_and_file(file_path, token, chat_id, agenda_date):
    application = ApplicationBuilder().token(token).build()
    async with application:
        try:
            message_text = f"📅 Berikut adalah agenda pimpinan untuk tanggal *{agenda_date}*."
            await application.bot.send_message(chat_id=chat_id, text=message_text, parse_mode="Markdown")
            print("✅ Pesan teks terkirim.")
        except Exception as e:
            print(f"⚠️ Gagal kirim pesan teks: {e}")
        try:
            with open(file_path, 'rb') as file:
                await application.bot.send_document(chat_id=chat_id, document=file, filename=os.path.basename(file_path))
            print("✅ PDF berhasil dikirim ke Telegram.")
        except Exception as e:
            print(f"❌ Gagal kirim file ke Telegram: {e}")

try:
    asyncio.run(send_telegram_message_and_file(pdf_file_name, bot_token, chat_id, tomorrow_str))
except Exception as e:
    print(f"❌ Terjadi kesalahan saat kirim ke Telegram: {e}")
