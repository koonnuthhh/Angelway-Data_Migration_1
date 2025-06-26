import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
import xlrd
from function.ColumnMappingFunction import map_excel_columns 
from function.replace_text_in_column import replace_text_in_column
from function import transfer_branchcode as transfer_branchcode

output_file = "Template_9_WL_output.xlsx"

def dowload_df(sourcepath, sheet_index=1):
   try:
      dowload_file = pd.read_excel(sourcepath, sheet_name=sheet_index)
      return dowload_file
   except Exception as e:
      print(f"The format is not xlsx.\n Changing read method to xls...")
      try:
         # เปิด workbook
         workbook = xlrd.open_workbook(sourcepath)
         # เลือก sheet ที่ต้องการ
         sheet = workbook.sheet_by_index(sheet_index)
         #อ่าน header
         headers = sheet.row_values(0)
         # อ่านข้อมูลที่เหลือ
         data = [sheet.row_values(row_idx) for row_idx in range(1, sheet.nrows)]
         # สร้าง DataFrame
         print("Dowload {sourcepath} success!!")
         
         return pd.DataFrame(data, columns=headers)
      except Exception as e:
         print(f"The format is not xls as well.\n Changing read method to csv(txt)...")
         df = pd.read_csv(sourcepath, delimiter='\t',  encoding='cp874')
         print("Dowload {sourcepath} success!!")
         return df
      
      
      
      
      
   
def Template_9_WL(source_file,b_zad_path,destination_file):
 source_sheet = "Sheet1"
 destination_sheet = "ข้อมูลการรับชำระ"

 column_mapping = {
    'เอกสาร': ['cont_no','cont_no.1'],
    'ลูกค้า': ['REF.','Ref_payment'],
    'เลขเอกสาร': 'payment_no',
    'ว/ท' : 'payment_date',
    'วิธีการชำระเงิน' : 'paytype_code',
    'ประเภทการชำระเงิน' : 'payfor_code',
    'หนี้ค้างชำระต่องวด' : 'payment',
    'ปี' : 'branch_code',
    'วันที่ฐาน' : 'effective_date',
 }

 template_df = map_excel_columns(source_file,destination_file,source_sheet,destination_sheet,column_mapping)
 
 bzad_pd = dowload_df(b_zad_path)
 print("Dowload Bzad success")
 template_df['cont_type'] = 'H'
 template_df['effective_date'] = template_df['effective_date'].fillna(template_df['payment_date'])
 template_df['net_payment'] = np.where(template_df['payfor_code'] == 3, template_df['payment'], template_df['payment']/1.07)
 template_df['vat_payment'] = np.where(template_df['payfor_code'] == 3, 0, template_df['net_payment'] *7/100)
 template_df['paytype_code'] = template_df['paytype_code'].astype(str).replace({
 '4': '003',
 '6': '009',
 '7': '010',
 '8': '012',
 '1': '101',
 '2': '201',
 '5': '202',
 'B': '601',
 'G': '602',
 '9': '603',
 '' : ''
 })
 template_df['payfor_code'] = template_df['payfor_code'].astype(str).replace({
 '1': '7001',
 '2': '1001',
 '3': '4001',
 '4': '2009',
 '' : ''
 })
 payment_df=transfer_branchcode.start(template_df,bzad_pd)
 template_df.rename(columns={'cont_no.1': 'cont_no'}, inplace=True)
 template_df['Ref_payment'] = payment_df['payment_no_with_year']
 template_df['branch_code'] = payment_df['branch_code']
 template_df.to_excel(output_file, index=False)
 print('🎉 Save success!')


# Template_9_WL(source_file = r"D:/Angelway/Migration to python/File_testing/WL_test/Tem.9/ZHP_PAYMENT.XLSX" ,b_zad_path = r"D:/Angelway/Migration to python/File_testing/WL_test/Tem.9/BSAD_WL.xlsx",destination_file = r"D:/Angelway/Migration to python/File_testing/WL_test/Tem.9/9-ข้อมูลการรับชำระ.xlsx")
