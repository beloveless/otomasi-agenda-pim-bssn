import gspread, datetime, requests, time
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, CellFormat, set_border, Color
from datetime import datetime as dt
import os

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds_path = os.path.join(os.path.dirname(__file__), 'teamup-425709-8333962f3c28.json')

creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)

spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
spreadsheet = client.open_by_key(spreadsheet_id)

today = datetime.date.today()
tomorow = today + datetime.timedelta(days=1)
tomorow_str = tomorow.strftime('%Y-%m-%d')
tomorow_sheet = tomorow.strftime("%d").lstrip('0')

worksheet_name = f'{tomorow_sheet}'
try:
    worksheet = spreadsheet.worksheet(worksheet_name)
except gspread.WorksheetNotFound:
    print(f"Worksheet dengan nama '{worksheet_name}' tidak ditemukan.")
    exit(1)

def redraw_default_blocks(ws):
    all_values = ws.get_all_values()
    if len(all_values) < 6:
        ws.update('A1:H5', [['']*8]*5)
        ws.update('D2', [['AGENDA KEGIATAN PIMPINAN BSSN']])
        ws.update('D3', [[f"Kamis, {tomorow.strftime('%d %B %Y')}"]])
        ws.update('A5:H5', [["KEGIATAN KE-", "KA", "WAKA", "SESTAMA", "DEPUTI 1", "DEPUTI 2", "DEPUTI 3", "DEPUTI 4"]])
        for i in range(4):
            base_row = 6 + i * 7
            ws.update(f"A{base_row}", str(i + 1))
            ws.update(f"B{base_row}:I{base_row + 6}", [[''] * 8 for _ in range(7)])
            ws.merge_cells(f"A{base_row}:A{base_row + 6}")

    # Bersihkan isi kolom B sampai I di baris-blok tertentu
    rows_to_clear = [6, 7, 13, 14, 20, 21, 27, 28]
    columns_to_clear = range(2, 9)
    for col in columns_to_clear:
        for row in rows_to_clear:
            ws.update_cell(row, col, "")
            time.sleep(0.1)

    # Tambahkan border pakai set_border (SOLID hitam)
    solid_black = Color(0, 0, 0)
    for i in range(6, 33):
        set_border(ws, f"B{i}:I{i}", style='SOLID', color=solid_black)
        time.sleep(0.1)

redraw_default_blocks(worksheet)

print("Blok agenda dibersihkan dan border default diperbarui.")