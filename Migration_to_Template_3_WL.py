from function.ColumnMappingFunction import map_excel_columns
from function.split_value_by_keyword import split_value_by_keyword
from function.replace_text_in_column import replace_text_in_column
import pandas as pd
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

def Migration_to_Template_3_WL(source_file, destination_file, source_sheet, destination_sheet):
    column_mapping = {
        'chassis_no': 'chassis_no',
        'engine_no': 'engine_no',
        'product_type_code': 'product_type_code',
        'reg_no' : 'reg_no',
        'brand_code' : 'brand_code',
        'model_code' : 'model_code_1',
        'sub_model_code' : 'sub_model_code_1',
        'color_code' : 'color_code_1',
        'rate_book' : 'rate_book'
    }

    template3_df = map_excel_columns(
        source_file,
        destination_file,
        source_sheet,
        destination_sheet,
        column_mapping
    )

    steps = [
        (["RED2","RED", "R-B", "BUS", "B-R", "B-S", "BUE", "R-S", "BLU", "W-R", "B-S_P", "R-G_P",
        "RED_P", "B-G", "BUB", "SLV", "S-B", "WBU", "W-B", "BLK", "R-B_P", "G-B_P",
        "2W-B", "SBW", "R-W", "BUR", "G-B", "S-R", "BRN", "P-W", "WHT", "BUW", "G-W",
        "GRN", "Y-W", "B-Y", "W-G", "Y-B", "RBU", "GNB", "GBU", "GRY", "BGO", "G-R",
        "O-B", "EBU", "BUG", "P-G", "B-P", "R-G", "YEL", "W-P", "W-Y", "B-W", "BRW",
        "WGO", "PKW", "RGO", "BUY", "R-S_P", "G-Y", "WPK", "W-S", "RBW", "O-W", "B-O",
        "BRS", "RBR", "BBR", "BBU", "S-Y"],'color_code'),
        (["TH", "3TH", "4TH", "2TH", "TH2", "TH1", "TH3", "TH4",
        "TH6", "5TH", "9TH", "6TH", "8TH", "7TH",'(C)'],'sub_model_code')
    ]

    split_value_by_keyword(template3_df,'sub_model_code_1' , steps , 'model_code')

    replace_text_in_column(template3_df,'sub_model_code','(C)','C')
    replace_text_in_column(template3_df,'color_code','RED2','RED')

    template3_df.to_excel(destination_file, index=False)
    print('🎉 Save success!')

# Example usage
# process_excel_file(
#     r"C:\Users\Asus\Desktop\Tem.3\ZBFMM.xlsx",
#     r"C:\Users\Asus\Desktop\Tem.3\3-ข้อมูลหลักประกัน - รถเล่ม&ทะเบียน..xlsx",
#     "Sheet1",
#     "ข้อมูลหลักประกันรถ"
# )
