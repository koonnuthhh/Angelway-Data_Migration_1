import pandas as pd
import sys
import os
from function.ColumnMappingFunction import map_excel_columns 
from function.Function_Plate_Provice.plate_code import prepare

output_file = "Template_3_WM_output.xlsx"

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
    prepare(template_df)
    print('extract province code from plate code success!')

    template_df.to_excel(output_file, index=False)
    print('🎉 Save success!')