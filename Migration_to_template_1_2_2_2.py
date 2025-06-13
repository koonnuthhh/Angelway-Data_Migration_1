import pandas as pd
from function.FunctionForX_1AndX_2AndX_3.clean_name import clean_name
from function.FunctionForX_1AndX_2AndX_3.map_column_WM_1 import map_column_WM
from function.FunctionForX_1AndX_2AndX_3.MappingCleanData import update_xlsx_with_clean_data
from function.FunctionForX_1AndX_2AndX_3.ChangeFormatMoney import ChangeFormatMoney
from function.FunctionForX_1AndX_2AndX_3.define_customer_type import define_customer_type
from function.FunctionForX_1AndX_2AndX_3.setWL_location import process_address_file

def Migration_to_template_1_2_2_2(
    source_file,
    source_sheet,
    template_1_path,
    template_1_sheet,
    template_2_path,
    template_2_sheet,
    full_name_column="ชื่อ-สกุล ลูกค้า",
    address_column="ที่อยู่"
):
    # 1. Clean customer name
    df_clean = clean_name(source_file, source_sheet, full_name_column)

    # 2. Define how to map cleaned names into template
    column_map = {
        'คำนำหน้า': 'คำนำหน้า',
        'ชื่อ': 'ชื่อ',
        'นามสกุล': 'สกุล'
    }

    # 3. Update Template 1 with cleaned names
    update_xlsx_with_clean_data(
        template_1_path,
        df_clean,
        column_map,
        sheet_name=template_1_sheet
    )

    # 4. Apply column mapping logic for WM
    map_column_WM(source_file, template_1_path)
    map_column_WM(source_file, template_2_path)

    # 5. Classify customer type
    define_customer_type(template_1_path, template_1_sheet, "คำนำหน้า", "ประเภทลูกหนี้")

    # 6. Convert income format
    ChangeFormatMoney(template_1_path, template_1_sheet, "รายได้ประจำ")

    # 7. Extract address fields from raw file
    maw = process_address_file(source_file, address_column)

    # 8. Define mapping for address fields
    column_location_map = {
        'ที่อยู่': 'ที่อยู่',
        'ซอย': 'ซอย',
        'ถนน': 'ถนน',
        'ตำบล': 'ตำบล',
        'อำเภอ': 'อำเภอ',
        'จังหวัด': 'จังหวัด',
        'รหัสไปรษณีย์': 'ปณ'
    }

    # 9. Update Template 2 with extracted location data
    update_xlsx_with_clean_data(
        template_2_path,
        maw,
        column_location_map,
        sheet_name=template_2_sheet
    )

# Example Usage
# process_wm_data(
#     source_file=r"C:\Users\Asus\Desktop\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\zfloan 60 04.2025.xlsx",
#     source_sheet="Sheet1",
#     template_1_path=r"C:\Users\Asus\Desktop\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\Template 1.2 - WM.xlsx",
#     template_1_sheet="Template1_WM",
#     template_2_path=r"C:\Users\Asus\Desktop\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\Template 2.2 - WM.xlsx",
#     template_2_sheet="Template2_WM"
# )
