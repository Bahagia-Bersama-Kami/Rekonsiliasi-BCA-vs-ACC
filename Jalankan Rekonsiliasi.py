import os
import shutil
import subprocess
import glob
import sys

dapur_dir = "Dapur"
required_dapur_files = ["1_AccCleaner.py", "2_BcaCleaner.py", "3_ProcessingData.py", "__init__.py"]
req_file_1 = "Acc.xls"
req_file_2 = "Bca.xlsx"

missing_items = []

if not os.path.exists(dapur_dir):
    missing_items.append(dapur_dir)
else:
    for f in required_dapur_files:
        file_path = os.path.join(dapur_dir, f)
        if not os.path.exists(file_path):
            missing_items.append(file_path)

if not os.path.exists(req_file_1):
    missing_items.append(req_file_1)

if not os.path.exists(req_file_2):
    missing_items.append(req_file_2)

if missing_items:
    print("--> Proses digagalkan. File atau folder berikut tidak ditemukan:")
    for item in missing_items:
        print(f"--> {item}")
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

shutil.move(req_file_1, os.path.join(dapur_dir, req_file_1))
shutil.move(req_file_2, os.path.join(dapur_dir, req_file_2))

scripts_to_run = ["1_AccCleaner.py", "2_BcaCleaner.py", "3_ProcessingData.py"]
current_dir = os.getcwd()
os.chdir(dapur_dir)

try:
    for script in scripts_to_run:
        print(f"--> Menjalankan {script}...")
        subprocess.run([sys.executable, script], check=True)
except subprocess.CalledProcessError:
    print(f"--> Terjadi kesalahan saat menjalankan {script}. Proses dihentikan.")
    os.chdir(current_dir)
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

os.chdir(current_dir)

hasil_file = os.path.join(dapur_dir, "Hasil_Rekonsiliasi.xlsx")
if os.path.exists(hasil_file):
    shutil.copy(hasil_file, "Hasil_Rekonsiliasi.xlsx")
    print("--> File Hasil_Rekonsiliasi.xlsx berhasil dibuat dan disalin.")
else:
    print("--> Gagal: File Hasil_Rekonsiliasi.xlsx tidak ditemukan setelah proses.")

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

print("--> Semua proses selesai dan folder Dapur telah dibersihkan dari file sampah.")