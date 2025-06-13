import pandas as pd
import sys
import os

# คำนวณ path ของตัวเอง (ไฟล์ปัจจุบัน)
current_dir = os.path.dirname(os.path.abspath(__file__))

# ถ้า branchcode_change.py อยู่ในโฟลเดอร์เดียวกับไฟล์นี้
sys.path.append(current_dir)

# Now import the module (filename without .py)
import branchcode_change as bc

def start(df, Transfer):
    # สร้าง branch_code ใน Transfer
    Transfer['branch_code'] = (Transfer['Clrng doc.'].astype(str)
        + pd.to_datetime(Transfer['การหักล้าง'], errors='coerce').dt.year.astype(str))

    # สร้าง branch_code ใน df (สมมุติ df มี 'branch_code' กับ 'payment_no')
    df['branch_code'] = df['payment_no'].astype(str) + df['branch_code'].astype(str)
    df['payment_no_with_year'] = df['branch_code']
    # Merge หารหัสเซกชัน (อันนี้คือ first mapping)
    merged = df.merge(
        Transfer[['branch_code', 'รหัสเซกชัน']],
        on='branch_code',
        how='left',
        suffixes=('', '_new')
    )

    # ถ้ามี 'รหัสเซกชัน' ให้แทนที่ branch_code
    merged['branch_code'] = merged['รหัสเซกชัน'].fillna(merged['branch_code']).astype(str)

    # ลบ column 'รหัสเซกชัน'
    merged.drop(columns=['รหัสเซกชัน'], inplace=True)

    # ส่งไป mapping ต่อใน branchcode_change
    df_final = bc.start(merged)
    Nut = pd.DataFrame({
    'branch_code': df_final['branch_code'],
    'payment_no_with_year': df['payment_no_with_year']
     })

    return Nut
