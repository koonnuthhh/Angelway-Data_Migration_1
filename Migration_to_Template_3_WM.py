import pandas as pd
import sys
import os
from function.ColumnMappingFunction import map_excel_columns 
from function.Function_Plate_Provice.plate_code import prepare
from function.clean_column_values import clean_column_names

output_file = "Template_3_WM_output.xlsx"
clean_column_values = {
    "MITSUBISHI": ["Mitsubihi", "Mitsubshi", "Misubishi", "Mitsubashi", "Mitzubishi"],
    "MG": ["Mg", "mg", "MGG", "MG.", "M-G"],
    "TOYOTA": ["Toyata", "Totota", "Toyoya", "Toyoya", "T0yota"],
    "ISUZU": ["Iszuzu", "Isuzu", "Isuzu", "Iszsu", "Izuzu"],
    "DFSK": ["DFK", "DFSKK", "D-FSK", "DSFK", "DFFSK"],
    "CHEVROLET": ["Chevorlet", "Chevy", "Chevrolat", "Chev", "Cheverlot"],
    "FORD": ["Forrd", "Fod", "Foed", "F0rd", "Frd"],
    "HONDA": ["Hondai", "Hando", "Hondo", "Hondar", "Honnda", "ฮอนด้า"],
    "D-MAX": ["Dmax", "D Max", "D_MAZ", "DMAX", "D-Maax"],
    "HINO": ["Hinoo", "Hiino", "Hinno", "HINO.", "Hino-"],
    "GPX": ["GPXX", "GP-PX", "G-PX", "GPPX", "GXP"],
    "NISSAN": ["Nisan", "Nissin", "Nissam", "Nissaan", "Nissn"],
    "KUBOTA": ["Kubotaa", "Cubota", "Kobota", "Kuboto", "Kuboata"],
    "MAZDA": ["Mazdaa", "Masda", "Mazd", "Mazta", "Mazdah"],
    "SUBARU": ["Suberu", "Subaro", "SubaRu", "Subauru", "Sbaru"],
    "YAMAHA": ["Yamha", "Yamama", "Yamhaa", "Yamah", "Yamhaa", "YAMAYA"],
    "VESPA": ["Vesbaa", "Vesspa", "Vspa", "Veespa", "Vesap"],
}
def keep_valid_float(val):
    try:
        # Strip leading/trailing spaces and normalize value
        val_str = str(val).strip().replace(",", "")
        float_val = float(val_str)
        return float_val if float_val != 0.0 else "-"
    except (ValueError, TypeError):
        return "-"

def Migration_to_Template_3_WM(source_file,destination_file,source_sheet,destination_sheet) :

    column_mapping = {
        'เลขที่สัญญา': 'เลขที่สัญญา',
        'หมายเลขถัง': 'chassis_no',
        'หมายเลขเครื่อง': 'engine_no',
        'ประเภทสินเชื่อ' : 'product_type_code',
        'เลขทะเบียน' : 'reg_no',
        'ปีทีจดทะเบียน' : 'reg_date',
        'ยี่ห้อ' : 'brand_code',
        'รุ่น' : 'model_code',
        'ราคาประเมินตามรุ่นรถ' : 'rate_book',
        'ชื่อ-สกุล ลูกค้า' : 'ownership',
    }

    template_df = map_excel_columns(source_file,destination_file,source_sheet,destination_sheet,column_mapping)
    template_df  = template_df[template_df['เลขที่สัญญา'].astype(str).str.startswith("48")]
    model_code_df = clean_column_names(template_df,"brand_code",column_mapping)
    template_df["brand_code"] = model_code_df["model_code"]
    template_df['reg_date'] = template_df['reg_date'].astype(str).apply(
    lambda x: x if x.isdigit() and 1900 <= int(x) <= 2100 else '-'
    )
    template_df.drop(columns=["model_code"], inplace=True)
    template_df['rate_book'] = template_df['rate_book'].apply(keep_valid_float)
    
    prepare(template_df)
    print('extract province code from plate code success!')

    template_df.to_excel(output_file, index=False)
    print('🎉 Save success!')