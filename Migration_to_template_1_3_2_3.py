import pandas as pd
from function.FunctionForX_1AndX_2AndX_3.CleanNameFunction import clean_name
from function.FunctionForX_1AndX_2AndX_3.UniversalMappingWL3 import transform_guarantor_data
from function.FunctionForX_1AndX_2AndX_3.MappingCleanData import update_xlsx_with_clean_data
from function.FunctionForX_1AndX_2AndX_3.define_customer_type import define_customer_type
from function.FunctionForX_1AndX_2AndX_3.DataframeHandle import update_xlsx_with_clean_names_openpyxl
from function.FunctionForX_1AndX_2AndX_3.setWL_location import process_address_file

def process_guarantor_data(
    source_file,
    source_sheet,  # <-- New required parameter
    template_1_path,
    template_1_sheet,
    template_2_path,
    template_2_sheet,
    name_column="ชื่อ-สกุล",
    address_column="ที่อยู่ลูกค้า"
):
    # 1. Map data into Template 1.3 and 2.3
    transform_guarantor_data(source_file, template_1_path)
    transform_guarantor_data(source_file, template_2_path)

    # 2. Clean names into DataFrame
    cleaned_names = clean_name(source_file, name_column, source_sheet)
    
    # Convert list of dicts to DataFrame
    df_clean = pd.DataFrame(cleaned_names)

    column_map = {
        'คำนำหน้า': 'คำนำหน้า',
        'ชื่อ': 'ชื่อ',
        'นามสกุล': 'สกุล'
    }

    # Convert DataFrame to list of dicts for update_xlsx_with_clean_names_openpyxl
    cleaned_data = df_clean.to_dict('records')

    update_xlsx_with_clean_names_openpyxl(
        template_1_path,
        cleaned_data,
        column_map,
        sheet_name=template_1_sheet
    )

    # 3. Define customer type from prefix
    define_customer_type(
        source_file=template_1_path,
        workbook=template_1_sheet,
        source_column='คำนำหน้า',
        target_column='ประเภทลูกหนี้'
    )

    # 4. Extract location data from Template 2
    df_location = process_address_file(template_2_path, address_column)

    column_location_map = {
        'ตำบล': 'ตำบล',
        'อำเภอ': 'อำเภอ',
        'จังหวัด': 'จังหวัด',
        'รหัสไปรษณีย์': 'ปณ'
    }

    update_xlsx_with_clean_data(
        template_2_path,
        df_location,
        column_location_map,
        sheet_name=template_2_sheet
    )
