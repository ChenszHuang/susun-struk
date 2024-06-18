#terlebih dahulu pip install pandas / pip install pyxlsb / xlsxwriter
import pandas as pd

def read_xlsb(filename):
    with pd.ExcelFile(filename, engine='pyxlsb') as xls:
        sheet_names = xls.sheet_names
        data = pd.DataFrame()
        for sheet_name in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            data = pd.concat([data, df], ignore_index=True)
        return data

def generate_nrff(data):
    if data.empty:
        return data
    
    # Mengisi nilai kosong pada kolom utama jika ada
    data['KS'] = data['KS'].fillna('00')
    data['JAMF'] = data['JAMF'].ffill()
    data['USRF'] = data['USRF'].fillna('Unknown')
    
    # Mengelompokkan data berdasarkan KS, TGFF, JAMF, dan USRF
    data['Group'] = data.groupby(['KS', 'TGFF', 'JAMF', 'USRF']).ngroup()
    
    # Menyusun nomor struk per grup
    nrff_dict = {}
    for group in data['Group'].unique():
        group_data = data[data['Group'] == group]
        if not group_data.empty:
            ks = group_data['KS'].iloc[0]
            if ks not in nrff_dict:
                nrff_dict[ks] = 1
            nrff = f"{ks}{nrff_dict[ks]:05}"  # Format nomor struk
            data.loc[data['Group'] == group, 'NRFF'] = nrff
            nrff_dict[ks] += 1
    
    # Menghapus kolom grup sementara
    data.drop(columns=['Group'], inplace=True)
    return data

def process_and_save(input_path, output_path):
    data = read_xlsb(input_path)
    data_with_nrff = generate_nrff(data)
    data_with_nrff.to_excel(output_path, index=False, engine='xlsxwriter')

# Path file input dan output
input_path = '2109BELUMSORT.xlsb'  
output_path = '2109sorted2.xlsx'

# Memproses dan menyimpan data
process_and_save(input_path, output_path)
