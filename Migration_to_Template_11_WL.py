import pandas as pd
import os
from openpyxl import load_workbook
from function.ColumnMappingFunction import map_excel_columns 

output_file = "Template_11_WL_output.xlsx"

def Template_11_WL(source_file,destination_file):
 source_sheet = "Sheet1"
 destination_sheet = "ข้อมูลการรับชำระ"

 column_mapping = {
    'ปี': 'year',
    'ว/ทเอกสาร': 'tax_date',
    'วันที่ฐาน': 'effective_date',
    'การกำหนด': 'cont_no',
    'ค่างวด': 'payment',
    'ก่อน VAT': 'net_payment',
    'ภาษี': 'vat_payment'
 }


 mapped_df = map_excel_columns(
    source_file,
    destination_file,
    source_sheet,
    destination_sheet,
    column_mapping
 )

 Nut = pd.read_excel(source_file, sheet_name=source_sheet, usecols='J,K', engine='openpyxl',dtype= str)
 Nut.columns = ['เลขที่สาขา', 'เลขที่ใบแจ้งหนี้']
 mapped_df['tax_no'] = Nut['เลขที่สาขา'].astype(str) + '-' + Nut['เลขที่ใบแจ้งหนี้'].astype(str)
 mapped_df.to_excel(output_file, index=False)
 print("Template_WL_11 ready to use!")