import pandas as pd
import openpyxl

def update_xlsx_with_clean_data(source_file, cleaned_df, column_map, sheet_name='Sheet1'):
    """
    file_path: path to Excel file (.xlsx)
    cleaned_df: DataFrame that contains cleaned name parts (e.g. คำนำหน้า, ชื่อจริง, นามสกุล)
    column_map: dict mapping cleaned_df column name -> Excel column name
                e.g. { 'คำนำหน้า': 'Prefix', 'ชื่อจริง': 'First Name', 'นามสกุล': 'Last Name' }
    sheet_name: name of sheet to update (default is 'Sheet1')
    """
    print("Start updating file with clean data")
    # โหลดไฟล์ Excel
    df_original = pd.read_excel(source_file, sheet_name=sheet_name)

    # ตรวจสอบว่า column ที่จะเขียนทับมีอยู่ใน Excel หรือไม่
    for col in column_map.values():
        if col not in df_original.columns:
            print("⚠️ Column '{col}' not found in Excel. Will create new column.")

    # เขียนค่าที่ cleaned แล้วไปยัง Excel columns ตาม map
    for clean_col, excel_col in column_map.items():
        df_original[excel_col] = cleaned_df[clean_col]
        print(f'✅ Copied "{clean_col}" → "{excel_col}"')

    # เขียนกลับไปยังไฟล์เดิม
    with pd.ExcelWriter(source_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_original.to_excel(writer, sheet_name=sheet_name, index=False)
        
    print(f'\n🎉 Done! File updated in place:\n{source_file}')