import pandas as pd
import numpy as np
from openpyxl import load_workbook

print("--> Membaca file Acc.xls")

df_raw = pd.read_excel('Acc.xls', header=None)
target_cols = ["Tanggal", "No. Sumber", "Keterangan", "Penambahan", "Pengurangan", "Saldo"]
num_cols = ["Penambahan", "Pengurangan", "Saldo"]

header_idx = -1
col_mapping = {}

for idx, row in df_raw.iterrows():
    row_values = [str(x).strip() for x in row.values if pd.notna(x)]
    match_count = sum(1 for col in target_cols if col in row_values)
    
    if match_count == len(target_cols):
        header_idx = idx
        for col_name in target_cols:
            for i, val in enumerate(row.values):
                if str(val).strip() == col_name:
                    col_mapping[col_name] = i
                    break
        break

if header_idx != -1:
    print("--> Header berhasil ditemukan, memfilter data")
    data_rows = []
    
    for idx in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        tanggal_val = row.iloc[col_mapping["Tanggal"]]
        
        if pd.notna(tanggal_val) and str(tanggal_val).strip() != "":
            row_dict = {}
            for col in target_cols:
                val = row.iloc[col_mapping[col]]
                if col in num_cols:
                    if isinstance(val, str):
                        val_clean = str(val).replace('.', '').replace(',', '.')
                        try:
                            val = float(val_clean)
                        except:
                            pass
                    else:
                        try:
                            val = float(val)
                        except:
                            pass
                row_dict[col] = val
            data_rows.append(row_dict)

    df_clean = pd.DataFrame(data_rows)
    
    for col in num_cols:
        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
        
    df_clean.replace(r'^\s*$', np.nan, regex=True, inplace=True)
    df_clean.dropna(subset=target_cols, inplace=True)

    df_clean.to_excel('Acc_temp.xlsx', index=False)
    print("--> Menyimpan data dan menerapkan autofit serta format angka")

    wb = load_workbook('Acc_temp.xlsx')
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        column_letter = col[0].column_letter
        col_header = col[0].value
        
        for cell in col:
            if cell.row > 1 and col_header in num_cols:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
            
            try:
                if cell.value:
                    if cell.row > 1 and col_header in num_cols and isinstance(cell.value, (int, float)):
                        formatted_str = f"{cell.value:,.2f}"
                        cell_length = len(formatted_str)
                    else:
                        cell_length = len(str(cell.value))
                        
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        ws.column_dimensions[column_letter].width = max_length + 2

    wb.save('Acc_temp.xlsx')
    print("--> Proses selesai, file Acc_temp.xlsx siap digunakan")
else:
    print("--> Kolom yang dicari tidak ditemukan di dalam file")