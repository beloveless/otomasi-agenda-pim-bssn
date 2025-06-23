import schedule
import time
import subprocess
import os

# Fungsi untuk menjalankan skrip malam
def run_script1():
    try:
        # Path relatif ke skrip target
        script_path = os.path.join(os.path.dirname(__file__), 'bot1.py')
        subprocess.run(['python', script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Failed to run script: {e}")
def run_script2():
    try:
        # Path relatif ke skrip target
        script_path = os.path.join(os.path.dirname(__file__), 'default.py')
        subprocess.run(['python', script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Failed to run script: {e}")


# Jadwalkan untuk menjalankan skrip setiap hari pada pukul 6:00 pagi
schedule.every().day.at("11:12").do(run_script1)
schedule.every().day.at("11:14").do(run_script2)

# Loop untuk menjalankan pending tasks
while True:
    schedule.run_pending()
    time.sleep(1)
