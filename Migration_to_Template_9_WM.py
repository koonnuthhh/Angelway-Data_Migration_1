import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import numbers
from datetime import datetime
from function.ColumnMappingFunction import map_excel_columns
import warnings
from function import mapping_for9

base_path = os.path.dirname(os.path.abspath(__file__))
destination_file = os.path.join(base_path, "Template", "9-ข้อมูลการรับชำระ_WM.xlsx")

# === 🔧 Example usage ===
def Template_9_WM (source_file,reference_file):
 source_sheet = "Sheet1"
 destination_sheet = "ข้อมูลการรับชำระ"

 column_mapping = { 
    'สาขา': 'branch_code',
    'เลขใบเสร็จ': 'payment_no',
    'ว/ทใบเสร็จ': ['payment_date','effective_date'],
    'Ref': 'REF.',
    'เลขที่สัญญา': 'cont_no',
    'ค่างวด': 'payment',
    'เงินต้น': 'principal_paid',
    'ดอกเบี้ย': 'interest_paid',
    # 'ประเภทการชำระ': 'ประเภทการชำระ'
 }
 mapped_df =map_excel_columns(source_file,destination_file,source_sheet,destination_sheet,column_mapping)
 mapped_df['total_payment'] = mapped_df['principal_paid'] + mapped_df['interest_paid']
 mapped_df['Diff (N & T)']= mapped_df['payment'] - mapped_df['total_payment']
 reference_df = pd.read_excel(reference_file)
 reference_df = mapping_for9.start(reference_df, 2)
 reference_df['paytype_code_x'] = reference_df['paytype_code']
 mapped_df = mapped_df.merge(
    reference_df[['เลขเอกสาร', 'paytype_code_x']].drop_duplicates('เลขเอกสาร'),
    left_on='payment_no',
    right_on='เลขเอกสาร',
    how='left'
 ).drop(columns='เลขเอกสาร')
 mapped_df['paytype_code'] = mapped_df['paytype_code_x']
 mapped_df['cont_type'] = 'E'
 mapped_df['payfor_code'] = '1001'
 mapped_df['discount']= 0.00
 mapped_df = mapped_df.drop(columns=['paytype_code_x'])
 mapped_df.to_excel(destination_file, index=False)
 print('🎉 Save success!')


