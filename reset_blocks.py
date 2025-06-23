import gspread
import time
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, CellFormat, Border

# Autentikasi
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds_path = 'teamup-425709-8333962f3c28.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
client = gspread.authorize(creds)

spreadsheet_id = '1vn6sMouwi9OOkgSdDNg18Hz_UzTsEFSvQH--WSpOHP4'
spreadsheet = client.open_by_key(spreadsheet_id)

# Tanggal besok
today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
sheet_name = str(int(tomorrow.strftime("%d")))

try:
    worksheet = spreadsheet.worksheet(sheet_name)
except:
    print(f"Worksheet dengan nama {sheet_name} tidak ditemukan.")
    exit(1)

def redraw_default_blocks(ws):
    # Header
    ws.batch_clear(['A1:H5'])
    ws.update(range_name='D2', values=[["AGENDA KEGIATAN PIMPINAN BSSN"]])
    ws.update(range_name='D3', values=[[f"Kamis, {tomorrow.strftime('%d %B %Y')}"]])
    ws.update(range_name='A5:H5', values=[["KEGIATAN KE-", "KA", "WAKA", "SESTAMA", "DEPUTI 1", "DEPUTI 2", "DEPUTI 3", "DEPUTI 4"]])

    # Gambar ulang blok kegiatan
    for i in range(4):
        start_row = 6 + i * 7
        ws.update(range_name=f"A{start_row}", values=[[str(i + 1)]])
        ws.update(range_name=f"B{start_row}:I{start_row + 6}", values=[[''] * 8 for _ in range(7)])
        ws.merge_cells(f"A{start_row}:A{start_row + 6}")

        # Tambah border
        border_fmt = CellFormat(
            borders=Border('SOLID', 'SOLID', 'SOLID', 'SOLID')
        )
        for r in range(start_row, start_row + 7):
            format_cell_range(ws, f"B{r}:I{r}", border_fmt)
            time.sleep(0.2)

redraw_default_blocks(worksheet)
print("✅ Blok agenda berhasil direset ke default.")
