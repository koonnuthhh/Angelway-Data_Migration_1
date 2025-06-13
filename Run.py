import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from Migration_to_Template_3_WL import Migration_to_Template_3_WL
from Migration_to_Template_3_WM import Migration_to_Template_3_WM
from Migration_to_Template_7_WL import Migration_to_Template_7_WL
from Migration_to_Template_6_WM import Migration_to_Template_6_WM

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