#!/usr/bin/env python
# coding: utf-8

# MASTER CODE

# In[1]:


import pandas as pd
import re
import math
import sys
import time
from datetime import datetime
from dateutil.parser import parse as dateutil_parse

# --- Function for Loading Bar ---
def update_progress(current, total):
    bar_length = 30  # length of progress bar
    filled = int(bar_length * current / total)
    bar = "█" * filled + "-" * (bar_length - filled)
    percent = (current / total) * 100
    sys.stdout.write(f"\rProcessing files: |{bar}| {percent:5.1f}% ({current}/{total})")
    sys.stdout.flush()

# --- Detect month name ---
def parse_excel_date(value):
    """
    Robustly parse Excel-style dates into dd/mm/YYYY strings.
    - Handles Excel serial numbers (ints/floats),
    - datetime objects,
    - ISO-like strings that start with a 4-digit year (parses year-first),
    - Indonesian month names (e.g., '12 Januari 2025'),
    - messy strings (removes 'nan'), and falls back to dateutil when needed.
    Returns: 'dd/mm/YYYY' or None if parsing failed.
    """
    if pd.isna(value):
        return None

    # Excel serial numbers (common when reading with pandas if value is numeric)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            # Excel's epoch (Excel for Windows): 1899-12-30
            dt = pd.to_datetime(value, unit='d', origin='1899-12-30')
            return dt.strftime('%m/%d/%Y')
        except Exception:
            pass

    # Already a datetime-like object
    if isinstance(value, datetime):
        return value.strftime('%m/%d/%Y')

    # Work with strings
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None

        # Remove stray 'nan' tokens
        s = re.sub(r'\s+nan\b', '', s, flags=re.IGNORECASE).strip()

        # If string starts with a 4-digit year (ISO-like), parse year-first explicitly
        if re.match(r'^\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}', s):
            # Try common year-first formats with/without time
            for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d', '%Y%m%d'):
                try:
                    dt = datetime.strptime(s, fmt)
                    return dt.strftime('%m/%d/%Y')
                except Exception:
                    continue
            # Last resort for year-first using dateutil but forcing year-first behavior:
            try:
                dt = dateutil_parse(s, yearfirst=True)
                return dt.strftime('%m/%d/%Y')
            except Exception:
                pass

        # Replace Indonesian month names with numbers to help parsing
        s_mod = s
        for bulan_name, bulan_num in daftar_bulan.items():
            # replace full month name with its numeric (so "12 Januari 2025" -> "12 1 2025")
            s_mod = re.sub(bulan_name, str(bulan_num), s_mod, flags=re.IGNORECASE)

        # Try parsing common day-first formats explicitly
        for fmt in ('%d %m %Y', '%d %m %y', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y',
                    '%d %m %Y %H:%M', '%d/%m/%Y %H:%M:%S'):
            try:
                dt = datetime.strptime(s_mod, fmt)
                return dt.strftime('%m/%d/%Y')
            except Exception:
                continue

        # Fallback to dateutil: try dayfirst=True first (Indonesian style),
        # then try without dayfirst if that fails.
        try:
            dt = dateutil_parse(s_mod, dayfirst=True)
            return dt.strftime('%m/%d/%Y')
        except Exception:
            try:
                dt = dateutil_parse(s_mod, dayfirst=False)
                return dt.strftime('%m/%d/%Y')
            except Exception:
                return None

    # If we reach here, parsing failed
    return None

daftar_bulan = {
    'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4, 'Mei': 5, 'Juni': 6,
    'Juli': 7, 'Agustus': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
}

# --- Create template dataframe similar with the master database ---
columns_temp_df = [
    'Tanggal Sampling', 'Nama Sampel', 'WHP (barg)', 'FCV (%)',
    'T Samp Brine', 'T Samp Steam', 'Samp Brine (barg)', 'Samp Steam (barg)',
    'P_sep (kscg)', 'Enthalpy (kJ/kg)', 'Flowrate Brine (kg/s)', 'Flowrate Steam (kg/s)',
    'Flowrate Brine (t/h)', 'Flowrate Steam (t/h)', 'TMF (t/h)', 
    
    'W-pH pada suhu 25°C', 'W-TDS kalkulasi*', 'W-Na+', 'W-K+', 'W-Ca2+', 'W-Mg2+', 'W-NH4',
    'W-Li+', 'W-Fe2+/3+', 'W-Al3+', 'W-F-', 'W-HCO3¯', 'W-Cl¯', 'W-SO42¯', 'W-B', 'W-SiO2',
    'W-As', 'W-H2S', 'W-CO2', 'W-Sr', 'W-Ba', 'W-Sb', 'W-Mn', 'W-2D', 'W-18O', 
    
    'C-pH pada suhu 25°C', 'C-TDS Kalkulasi*', 'C-Na+', 'C-K+', 'C-Ca2+', 'C-Mg2+', 'C-NH4',
    'C-Li+', 'C-Fe2+/3+', 'C-Al3+', 'C-F-', 'C-HCO3¯', 'C-Cl¯', 'C-SO42¯', 'C-B', 'C-SiO2', 
    'C-As', 'C-H2S', 'C-CO2', 'C-Sr', 'C-Ba', 'C-Hg', 'C-Mn', 'C-2D', 'C-18O', 
    
    'Total NCGs (%wt)', 'g-CO2', 'g-H2S', 'g-NH3', 'g-Ar', 'g-N2', 'g-CH4', 'g-H2', 'Air Cont.',
    'g(ppm)-CO2', 'g(ppm)-H2S', 'g(ppm)-NH3', 'g(ppm)-He', 'g(ppm)-H2', 'g(ppm)-Ar',
    'g(ppm)-N2', 'g(ppm)-CH4'
]


while True:
    temp_df = pd.DataFrame(columns=columns_temp_df)
    
    print('''
____________________________
3  S  T  E  P  S  C  H  E  M
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
HELLO GEOCHEMIST! :D
Welcome to 3STEPSCHEM, a converter made by Alvin & Ifan to swiftly
convert your lab result into ready-to-use excel format, may it
help you with your daily work!


    ''')
    
    while True:
        file = []
        coupon = 1
        total_sheets = 0
        progress_count = 1
        
        # --- Input every files needed ---
        while coupon == 1:
            user_input = input("Enter file path or filename (e.g., SPW_Agust.xlsx): ")
            file.append(user_input)
            total_sheets += len(pd.ExcelFile(user_input).sheet_names[1:])
            coupon = int(input("Add more file? 1=Yes | 0=No (type only the number): "))
    
        # --- Processing! ---
        print('\n')    
        for f in file:
            sheet_list = pd.ExcelFile(f).sheet_names[1:]
            for sheet in sheet_list:
                update_progress(progress_count, total_sheets)
                
                # --- Input Lab SPW/SCS/NCG Data ---
                file_path = f
                sheet_proc = sheet
                    
                df_raw = pd.read_excel(file_path, sheet_name=sheet_proc, header=None)
                df_raw = df_raw.dropna(axis=1, how='all')
                for col in df_raw.columns:
                    df_raw[col] = df_raw[col].astype(str).str.replace(r'<\s*', '', regex=True)
                    df_raw[col] = df_raw[col].astype(str).str.replace(',', '.', regex=False)
                    
                # --- Extract Data Identity ---
                header_row_index = None
                nama_value = None
                tgl = None
                jenis_value = None
                if 'NCG' in file_path:
                    for i, row in df_raw.iterrows():
                        row_str = " ".join(row.astype(str).tolist())  # convert entire row to single string
                        if 'NAMA SAMPEL' in row_str.upper():
                        # --- SCENARIO 1 ---      
                            if '\n' in row_str.upper():
                                # extract text from 'NAMA SAMPEL' until next newline or end of string
                                nama_sampel = re.search(r'NAMA SAMPEL.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if nama_sampel:
                                    nama = nama_sampel.group(0).strip()
                                    # extract text after ':' if present
                                    nama_match = re.search(r':\s*(.*)', nama)
                                    if nama_match:
                                        nama_value = nama_match.group(1).strip()
                                
                                tanggal_sampling = re.search(r'TANGGAL SAMPLING.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if tanggal_sampling:
                                    tanggal = tanggal_sampling.group(0).strip()
                                    # extract text after ':' if present
                                    tanggal_match = re.search(r':\s*(.*)', tanggal)
                                    if tanggal_match:
                                        tanggal_value = tanggal_match.group(1).strip()
                                        tgl = parse_excel_date(tanggal_value)
                                
                                jenis_sampel = re.search(r'JENIS SAMPEL.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if jenis_sampel:
                                    jenis = jenis_sampel.group(0).strip()
                                    # extract text after ':' if present
                                    jenis_match = re.search(r':\s*(.*)', jenis)
                                    if jenis_match:
                                        jenis_value = jenis_match.group(1).strip()
                                continue
                        # --- SCENARIO 2 ---
                            else:
                                nama_sampel = re.search(r'NAMA SAMPEL\s*(.*)', row_str)
                                if nama_sampel:
                                    nama = nama_sampel.group(1).strip()
                                    nama_value = re.sub(r'\s+nan', '', nama, flags=re.IGNORECASE).strip()
                        
                        if 'TANGGAL SAMPLING' in row_str.upper():
                            tanggal_match = re.search(r'TANGGAL SAMPLING\s*(.*)', row_str)
                            if tanggal_match:
                                tanggal_value = tanggal_match.group(1).strip()
                                tanggal_value = re.sub(r'\s+nan', '', tanggal_value, flags=re.IGNORECASE).strip()
                                tgl = parse_excel_date(tanggal_value)
                        
                        if 'JENIS SAMPEL' in row_str.upper():
                            jenis_sampel = re.search(r'JENIS SAMPEL\s*(.*)', row_str)
                            if jenis_sampel:
                                jenis = jenis_sampel.group(1).strip()
                                jenis_value = re.sub(r'\s+nan', '', jenis, flags=re.IGNORECASE).strip()
                    
                    # --- Set Data Range ---    
                        if row.astype(str).str.contains('PARAMETER ANALISIS', case=False, na=False).any():
                            header_row_index = i+1
                            break
                                
                    # --- Extract Bulk NCG & Air Cont. ---
                    df_raw = df_raw.dropna(axis=1, how='all')
                    bulk_ncg = float(df_raw[df_raw[df_raw.columns[0]] == 'Persen Berat NCG'][df_raw.columns[1]].iloc[0])
                    air_cont = float(df_raw[df_raw[df_raw.columns[3]] == 'Persen udara dalam sampel'][df_raw.columns[4]].iloc[0])
                        
                    # --- Set Index Column ---
                    if header_row_index is not None:
                        df = pd.read_excel(file_path, sheet_name=sheet_proc, header=header_row_index)
                    else:
                        print("No row containing 'ANALISIS' was found.")
                    # --- Data Cleansing ---
                    for i in range(len(df)):
                        if math.isnan(df['% Mol Gas'][i]):
                            df = df.drop(df.index[i:])
                            break
                    df = df.rename(columns={df.columns[0]: 'PARAMETER ANALISIS'})
                    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed', case=False)]
                    for cols in df.columns[1:]:
                        df[cols] = df[cols].astype(str).str.replace(r'<\s*', '', regex=True)
                        df[cols] = df[cols].astype(str).str.replace(',', '.', regex=False)
                        df[cols] = pd.to_numeric(df[cols], errors='coerce')
                
                else:
                    for i, row in df_raw.iterrows():
                        row_str = " ".join(row.astype(str).tolist())  # convert entire row to single string
                        if 'NAMA SAMPEL' in row_str.upper():
                        # --- SCENARIO 1 ---      
                            if '\n' in row_str.upper():
                                # extract text from 'NAMA SAMPEL' until next newline or end of string
                                nama_sampel = re.search(r'NAMA SAMPEL.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if nama_sampel:
                                    nama = nama_sampel.group(0).strip()
                                    # extract text after ':' if present
                                    nama_match = re.search(r':\s*(.*)', nama)
                                    if nama_match:
                                        nama_value = nama_match.group(1).strip()
                                
                                tanggal_sampling = re.search(r'TANGGAL SAMPLING.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if tanggal_sampling:
                                    tanggal = tanggal_sampling.group(0).strip()
                                    # extract text after ':' if present
                                    tanggal_match = re.search(r':\s*(.*)', tanggal)
                                    if tanggal_match:
                                        tanggal_value = tanggal_match.group(1).strip()
                                        tgl = parse_excel_date(tanggal_value)
                                
                                jenis_sampel = re.search(r'JENIS SAMPEL.*?(?=\n|$)', row_str, flags=re.IGNORECASE)
                                if jenis_sampel:
                                    jenis = jenis_sampel.group(0).strip()
                                    # extract text after ':' if present
                                    jenis_match = re.search(r':\s*(.*)', jenis)
                                    if jenis_match:
                                        jenis_value = jenis_match.group(1).strip()
                                continue
                        # --- SCENARIO 2 ---
                            else:
                                nama_sampel = re.search(r':\s*(.*)', row_str)
                                if nama_sampel:
                                    nama = nama_sampel.group(1).strip()
                                    nama_value = re.sub(r'\s+nan', '', nama, flags=re.IGNORECASE).strip()
                        
                        if 'TANGGAL SAMPLING' in row_str.upper():
                            tanggal_match = re.search(r':\s*(.*)', row_str)
                            if tanggal_match:
                                tanggal_value = tanggal_match.group(1).strip()
                                tanggal_value = re.sub(r'\s+nan', '', tanggal_value, flags=re.IGNORECASE).strip()
                                tgl = parse_excel_date(tanggal_value)
                        
                        if 'JENIS SAMPEL' in row_str.upper():
                            jenis_sampel = re.search(r':\s*(.*)', row_str)
                            if jenis_sampel:
                                jenis = jenis_sampel.group(1).strip()
                                jenis_value = re.sub(r'\s+nan', '', jenis, flags=re.IGNORECASE).strip()
                        
                    # --- Set Data Range ---    
                        if row.astype(str).str.contains('PARAMETER ANALISIS', case=False, na=False).any():
                            header_row_index = i
                            break
                    # --- Set Index Column ---
                    if header_row_index is not None:
                        df = pd.read_excel(file_path, sheet_name=sheet_proc, header=header_row_index)
                        
                        if 'NO' in df.columns:
                            df = df.set_index('NO')
                        else:
                            print("Column 'NO' not found; index not set.")
                        
                    else:
                        print("No row containing 'ANALISIS' was found.")
                    # --- Data Cleansing ---
                    for i in range(len(df)):
                        if type(df.index[i]) != int:
                            df = df.drop(df.index[i:])
                            break
                    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed', case=False)]
                    df['HASIL'] = df['HASIL'].astype(str).str.replace(r'<\s*', '', regex=True)
                    df['HASIL'] = df['HASIL'].astype(str).str.replace(',', '.', regex=False)
                    df['HASIL'] = pd.to_numeric(df['HASIL'], errors='coerce')
                    # df
                
                # --- Insert lab result to the master database format ---
                if tgl not in temp_df['Tanggal Sampling'].values:
                    temp_df.loc[len(temp_df), 'Tanggal Sampling'] = tgl
                    temp_df.loc[len(temp_df)-1, 'Nama Sampel'] = nama_value
                    if jenis_value == 'SPW':
                        for i in range(1, len(df)+1):
                            for item in temp_df.columns[15:38]:
                                if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                    temp_df.loc[len(temp_df)-1, item] = df.loc[i]['HASIL']
                                    break
                    elif jenis_value == 'SCS':
                        for i in range(1, len(df)+1):
                            for item in temp_df.columns[40:63]:
                                if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                    temp_df.loc[len(temp_df)-1, item] = df.loc[i]['HASIL']
                                    break
                    elif jenis_value == 'GAS':
                        for i in range(0, len(df)):
                            for item in temp_df.columns[66:73]:
                                if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                    temp_df.loc[len(temp_df)-1, item] = df.loc[i]['% Mol Gas']
                                    break
                            for item in temp_df.columns[74:82]:
                                if item[7:] in df.loc[i]['PARAMETER ANALISIS']:
                                    temp_df.loc[len(temp_df)-1, item] = df.loc[i]['ppmw']
                                    break
                            temp_df.loc[len(temp_df)-1, 'Total NCGs (%wt)'] = bulk_ncg
                            temp_df.loc[len(temp_df)-1, 'Air Cont.'] = air_cont
                            
                else:
                    duplo_loc = None
                    for n in range(0, len(temp_df)):
                        if temp_df.iloc[n]['Tanggal Sampling'] == tgl and temp_df.iloc[n]['Nama Sampel'] == nama_value:
                            duplo_loc = n
                            break
                            
                    if duplo_loc != None:
                        if jenis_value == 'SPW':
                            for i in range(1, len(df)+1):
                                for item in temp_df.columns[15:38]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        # temp_df.loc[temp_df['Tanggal Sampling'] == tgl, item] = df.loc[i]['HASIL']
                                        temp_df.loc[duplo_loc, item] = df.loc[i]['HASIL']
                                        break
                        elif jenis_value == 'SCS':
                            for i in range(1, len(df)+1):
                                for item in temp_df.columns[40:63]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        # temp_df.loc[temp_df['Tanggal Sampling'] == tgl, item] = df.loc[i]['HASIL']
                                        temp_df.loc[duplo_loc, item] = df.loc[i]['HASIL']
                                        break
                        elif jenis_value == 'GAS':
                            for i in range(0, len(df)):
                                for item in temp_df.columns[66:73]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[duplo_loc, item] = df.loc[i]['% Mol Gas']
                                        break
                                for item in temp_df.columns[74:82]:
                                    if item[7:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[duplo_loc, item] = df.loc[i]['ppmw']
                                        break
                                temp_df.loc[duplo_loc, 'Total NCGs (%wt)'] = bulk_ncg
                                temp_df.loc[duplo_loc, 'Air Cont.'] = air_cont
                
                    else:
                        temp_df.loc[len(temp_df), 'Tanggal Sampling'] = tgl
                        temp_df.loc[len(temp_df)-1, 'Nama Sampel'] = nama_value
                        if jenis_value == 'SPW':
                            for i in range(1, len(df)+1):
                                for item in temp_df.columns[15:38]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[len(temp_df)-1, item] = df.loc[i]['HASIL']
                                        break
                        elif jenis_value == 'SCS':
                            for i in range(1, len(df)+1):
                                for item in temp_df.columns[40:63]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[len(temp_df)-1, item] = df.loc[i]['HASIL']
                                        break
                        elif jenis_value == 'GAS':
                            for i in range(0, len(df)):
                                for item in temp_df.columns[66:73]:
                                    if item[2:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[len(temp_df)-1, item] = df.loc[i]['% Mol Gas']
                                        break
                                for item in temp_df.columns[74:82]:
                                    if item[7:] in df.loc[i]['PARAMETER ANALISIS']:
                                        temp_df.loc[len(temp_df)-1, item] = df.loc[i]['ppmw']
                                        break
                                temp_df.loc[len(temp_df)-1, 'Total NCGs (%wt)'] = bulk_ncg
                                temp_df.loc[len(temp_df)-1, 'Air Cont.'] = air_cont
                progress_count += 1
    
        end_processing = input('\nProcess more files? (y/n): ')
        if end_processing == 'n':
            break
    
    # --- Export temp_df to Excel ---
    output_path = input("\nInsert output file name (e.g., final_output.xlsx): ")   # change filename if needed
    temp_df.to_excel(output_path, index=False)
    print(f"DataFrame exported to: {output_path}")
    print('\nDone ;)')

    end_loop = input('\nStart again? (y/n): ')
    if end_loop == 'n':
        print('Quitting program...')
        break


# In[ ]:




