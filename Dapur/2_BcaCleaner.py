import pandas as pd
from openpyxl import load_workbook

print("--> Membaca file Bca.xlsx")

df_raw = pd.read_excel('Bca.xlsx', header=None)
target_cols = ["Tanggal Transaksi", "Keterangan", "Cabang", "Jumlah", "Saldo"]
out_cols = ["Tanggal Transaksi", "Keterangan", "Cabang", "Debet", "Kredit", "Saldo"]
num_cols = ["Debet", "Kredit", "Saldo"]

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
    print("--> Header berhasil ditemukan, memfilter dan memecah data")
    data_rows = []
    
    for idx in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        tanggal_val = row.iloc[col_mapping["Tanggal Transaksi"]]
        
        if pd.notna(tanggal_val) and str(tanggal_val).strip() != "":
            row_dict = {}
            
            try:
                date_obj = pd.to_datetime(tanggal_val, dayfirst=True)
                row_dict["Tanggal Transaksi"] = date_obj.strftime('%d/%m/%Y')
            except:
                row_dict["Tanggal Transaksi"] = tanggal_val
                
            row_dict["Keterangan"] = row.iloc[col_mapping["Keterangan"]]
            row_dict["Cabang"] = row.iloc[col_mapping["Cabang"]]
            
            jumlah_val = str(row.iloc[col_mapping["Jumlah"]]).strip()
            debet_val = None
            kredit_val = None
            
            if jumlah_val.endswith("DB"):
                val_clean = jumlah_val.replace("DB", "").strip().replace(',', '')
                try:
                    debet_val = float(val_clean)
                except:
                    pass
            elif jumlah_val.endswith("CR"):
                val_clean = jumlah_val.replace("CR", "").strip().replace(',', '')
                try:
                    kredit_val = float(val_clean)
                except:
                    pass
            
            row_dict["Debet"] = debet_val
            row_dict["Kredit"] = kredit_val
            
            saldo_val = row.iloc[col_mapping["Saldo"]]
            if isinstance(saldo_val, str):
                saldo_clean = saldo_val.replace(',', '')
                try:
                    saldo_val = float(saldo_clean)
                except:
                    pass
            else:
                try:
                    saldo_val = float(saldo_val)
                except:
                    pass
            row_dict["Saldo"] = saldo_val
            
            data_rows.append(row_dict)

    df_clean = pd.DataFrame(data_rows, columns=out_cols)
    df_clean.to_excel('Bca_temp.xlsx', index=False)
    print("--> Menyimpan data dan menerapkan autofit serta format angka dan tanggal")

    wb = load_workbook('Bca_temp.xlsx')
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

    wb.save('Bca_temp.xlsx')
    print("--> Proses selesai, file Bca_temp.xlsx siap digunakan")
else:
    print("--> Kolom yang dicari tidak ditemukan di dalam file")