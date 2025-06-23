import datetime

# Tanggal hari ini
today = datetime.date.today()

# Tanggal kemarin
yesterday = today - datetime.timedelta(days=1)

today_str = today.strftime("%d").lstrip('0')
yesterday_str = yesterday.strftime("%d").lstrip('0')

print("Tanggal hari ini:", today_str)
print("Tanggal kemarin:", yesterday_str)
