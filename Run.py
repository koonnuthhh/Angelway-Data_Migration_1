import sys
import os
import openpyxl
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from Migration_to_template_1_1_2_1 import Migration_to_template_1_1_2_1
from Migration_to_template_1_2_2_2 import Migration_to_template_1_2_2_2
from Migration_to_template_1_3_2_3 import Migration_to_template_1_3_2_3

from Migration_to_Template_3_WL import Migration_to_Template_3_WL
from Migration_to_Template_7_WL import Migration_to_Template_7_WL
from Migration_to_Template_9_WL import  Template_9_WL
from Migration_to_Template_11_WL import Template_11_WL

from Migration_to_Template_3_WM import Migration_to_Template_3_WM
from Migration_to_Template_6_WM import Migration_to_Template_6_WM
from Migration_to_Template_9_WM import Template_9_WM
from Migration_to_template_12_WM import Migration_to_template_12_WM

# print(openpyxl.__name__)
# print(type(openpyxl.__name__))
# ##Temp3_WL
# source_file = r"C:\Users\Asus\Desktop\Tem.3\ZBFMM.xlsx"
# destination_file = r"C:\Users\Asus\Desktop\Tem.3\3-ข้อมูลหลักประกัน - รถเล่ม&ทะเบียน..xlsx"
# source_sheet = "Sheet1"
# destination_sheet = "ข้อมูลหลักประกันรถ"

# Migration_to_Template_3_WL(source_file,destination_file,source_sheet,destination_sheet)

# ##Temp3_WM
# source_file = r"C:\Users\Asus\Desktop\Tem.3_WM\zfloan50 ใหม่.xlsx"
# destination_file = r"C:\Users\Asus\Desktop\Tem.3_WM\3-ข้อมูลหลักประกัน - รถเล่ม_ทะเบียน WM.xlsx"
# source_sheet = "Sheet1"
# destination_sheet = "ข้อมูลหลักประกันรถ"

# Migration_to_Template_3_WM(source_file,destination_file,source_sheet,destination_sheet)

# ##Temp7
# source_file_1 = r"C:\Users\andre\Desktop\Internship Data cleaning\Tem.7\wl_zbfamt_04.2025(Temp7.1).xlsx"
# source_file_2=r"C:\Users\andre\Desktop\Internship Data cleaning\Tem.7\wl_zbfamt2_04.2025 (Temp7.2).xlsx"
# destination_sheet = "ข้อมูลสัญญาเชื้อ"


# Declare variable
# ไฟล์ต้นฉบับ
# source_file1 = "WM_1014100 - 1014108_04.2025.xls" # WM_1014100
# source_file3 =  "zfloan50_04.2025.xlsx" # zfloan 50 / มี sheet เดียว
# zfloan_raw = "zfloan20_04.2025 ไฟล์ดิบ.txt"
# source_file4 = "zfloan 60 04.2025.xlsx" # zfloan 60 / มี sheet เดียว

# ## File temp 6
# destination_file = r"C:\InternJob\WM-Temp6\Temp6 - Copy.xlsx"

# Migration_to_Template_6_WM(source_file1,zfloan_raw,source_file3,source_file4,destination_file)

## File temp 12
# destination_file = "12-ข้อมูลค่างวดค้างชำระสัญญาเงินกู้(แยกเงินต้นดอกเบี้ย).xlsx"
# Migration_to_template_12_WM(destination_file)


## File temp 1.1,2.1
# source_file = "D:\Angelway\Migration to python\File_testing\WL_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WL\wl_zbicust_04.2025.xlsx"
# source_sheet = "Sheet1"
# template_1_path = "D:\Angelway\Migration to python\File_testing\WL_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WL\Template 1.1 - WL.xlsx"
# template_1_sheet = "Template1_WL"
# template_2_path = "D:\Angelway\Migration to python\File_testing\WL_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WL\Template 2.1 - WL.xlsx"
# template_2_sheet = "Template2_WL"

# Migration_to_template_1_1_2_1(source_file,source_sheet,template_1_path,template_1_sheet,template_2_path,template_2_sheet)

## File temp 1.2,2.2
# source_file = "D:\Angelway\Migration to python\File_testing\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\zfloan 60 04.2025.XLSX"
# source_sheet = "Sheet1"
# template_1_path = "D:\Angelway\Migration to python\File_testing\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\Template 1.2 - WM.xlsx"
# template_1_sheet = "Template1_WM"
# template_2_path = "D:\Angelway\Migration to python\File_testing\WM_test\ข้อมูลลูกค้า+ที่อยู่ลูกค้า WM\Template 2.2 - WM.xlsx"
# template_2_sheet = "Template2_WM"
# full_name_column="ชื่อ-สกุล ลูกค้า"
# address_column="ที่อยู่"

# Migration_to_template_1_2_2_2(source_file,source_sheet,template_1_path,template_1_sheet,template_2_path,template_2_sheet,full_name_column,address_column)