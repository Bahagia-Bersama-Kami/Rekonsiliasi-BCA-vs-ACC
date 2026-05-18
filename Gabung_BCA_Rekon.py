import os
import re
import pandas as pd

def get_sort_key(filename):
    match = re.search(r'BCA \d+ (\d+)', filename)
    if match:
        return int(match.group(1))
    return 999

file_list = [f for f in os.listdir('.') if re.match(r'BCA \d+', f) and f.endswith('.xlsx')]
file_list.sort(key=get_sort_key)

data_frames = []

for file_name in file_list:
    print("--> Memproses file: " + file_name)
    df = pd.read_excel(file_name, skiprows=5, usecols="A:E")
    df = df.dropna(how='any')
    data_frames.append(df)

if data_frames:
    merged_data = pd.concat(data_frames, ignore_index=True)
    merged_data.columns = ['Tanggal Transaksi', 'Keterangan', 'Cabang', 'Jumlah', 'Saldo']
    merged_data.to_excel('Bca.xlsx', index=False)
    print("--> Proses selesai. File Bca.xlsx berhasil dibuat.")
else:
    print("--> Tidak ada file atau data yang memenuhi kriteria.")
