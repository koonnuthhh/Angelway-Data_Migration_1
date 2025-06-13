import pandas as pd
from openpyxl import load_workbook
import numpy as np
import time
from function.WM_clean_zfloan20 import WM_clean_zfloan20

start_time = time.time()

def Migration_to_Template_6_WM(source_file1,zfloan_raw,source_file3,source_file4,destination_file) :
    
    #Clean zfloan20 first
    WM_clean_zfloan20(zfloan_raw)
    
    # Declare variable
    source_file2 = "Temp6_clean_excel\WM_Temp6_cleaned.xlsxx" # zfloan20 กรอง 1 
    source_file2_2 = "Temp6_clean_excel\WM_Temp6_cleanedV2.xlsx" # zfloan20 กรอง 2

    # หน้า sheet ของไฟล์ต้นฉบับ
    source_sheet1 = "รายละเอียด" # WM_1014100

    target_sheet = "ข้อมูลสัญญาลดต้นลดดอก" 

    # ชื่อ column เลขที่สัญญา ของไฟล์ engine
    id_column = "เลขที่สัญญา"

    ## ตัวแปรสำหรับกรอง zf20 ดิบ
    # กรองรอบ 1
    doc_col = "เลขเอกสาร" # reference เอกสาร
    clrng_col = "Clrng doc." # ใบแจ้งหนี้
    date_col = "การหักล้าง" # วันที่
    # กรองรอบ 2
    date_col_2 = "Pstng Date" # column วันที่สำหรับกรองรอบ 2
    date = "30/04/2025" # วันที่ที่จะกรองออก

    ## ตัวแปรสำหรับ zf60

    WM = pd.read_excel(source_file1, sheet_name=source_sheet1, engine="xlrd")

    print(WM.columns.tolist())

    WM = WM[["เลขที่สัญญา","สาขา","เงินต้นคงเหลือ"]].copy()

    WM = WM.rename(columns={
        'สาขา': 'branch_code',
        'เงินต้นคงเหลือ': 'current_balance'
    })  

    ## load data
    zf20_1 = pd.read_excel(source_file2, engine="openpyxl")

    zf20_1 = zf20_1.rename(columns={'เลขที่สัญญ': 'เลขที่สัญญา'})

    zf20_1 = zf20_1[["เลขที่สัญญา","        ดบ.ในงวด","วันครบกำหน"]].copy()

    zf20_1 = zf20_1.rename(columns={
        '        ดบ.ในงวด': 'interest_accrued',
        'วันครบกำหน': 'next_due_date'
    })

    ## load data
    zf20_2 = pd.read_excel(source_file2_2, engine="openpyxl")

    zf20_2 = zf20_2.rename(columns={'เลขที่สัญญ': 'เลขที่สัญญา'})

    zf20_2 = zf20_2[["เลขที่สัญญา","              เงินต้น","        ดบ.ในงวด",
                                "          มูลค่าในงวด","@วันในงวด"]].copy()

    zf20_2 = zf20_2.rename(columns={
        '              เงินต้น': 'installment_current',
        '        ดบ.ในงวด': 'installment_interest',
        '          มูลค่าในงวด': 'installment_over_due',
        '@วันในงวด': 'dpd'
    })

    ## load data
    zf60 = pd.read_excel(source_file4, engine="openpyxl")

    zf60['วันเริ่มต้น'] = pd.to_datetime(zf60['วันเริ่มต้น'], errors='coerce')

    zf60['payment_day'] = zf60['วันเริ่มต้น'].dt.day

    zf60['สถานะสัญญา'].astype

    # Example: Clean column 'col' to keep only leading digits as string

    zf60['สถานะสัญญา'] = zf60['สถานะสัญญา'].str.extract(r'ค้างชำระ\s+(\d+)')[0]
    zf60['สถานะสัญญา'] = pd.to_numeric(zf60['สถานะสัญญา'], errors='coerce').astype('Int64')

    zf60['สถานะสัญญา'] = zf60['สถานะสัญญา'].astype(str).replace('<NA>', '')

    # Mapping dictionary from description (คำอธิบาย) to loan type code (รหัสประเภทสินเชื่อ)
    loan_type_map = {
        "จำนำ-รถจยย.": "LN01",
        "จำนำ-รถจยย. พิเศษ": "LN02",
        "จำนำ-รถบรรทุก": "LT01",
        "จำนำ-รถยนต์": "LC01",
        "จำนำ-รถยนต์ พิเศษ": "LC02",
        "ผู้ประกอบอาชีพ": "NN01",
        "ส่วนบุคคล": "PL01",
        "ส่วนบุคคล โปรน้ำท่วม": "PL02",
        "ส่วนบุคคลโปร1": "PL03",
        "จำนำ-รถยนต์ ดบ.20%": "LC03"
    }

    interest_flat_rate_map = {
        "LN01": "24",
        "LN02": "18",
        "LT01": "18",
        "LC01": "24",
        "LC02": "18",
        "NN01": "33",
        "PL01": "25",
        "PL02": "8",
        "PL03": "23",
        "LC03": "20"
    }

    # map descriptions to loan type codes
    zf60["ประเภทสินเชื่อ"] = zf60["ประเภทสินเชื่อ"].map(loan_type_map)

    # replace value in int.rte with ประเภทสินเชื่อ
    zf60['int.rte'] = zf60['ประเภทสินเชื่อ']

    # map interest flat rate base on ประเภทสินเชื่อ
    zf60["int.rte"] = zf60["int.rte"].map(interest_flat_rate_map)

    zf60 = zf60[["สัญญา","ประเภทสินเชื่อ","int.rte","วันเริ่มต้น","วันสิ้นสุด","payment_day","สถานะสัญญา"]].copy()

    zf60.rename(columns={'สัญญา': 'เลขที่สัญญา'}, inplace=True)

    zf60 = zf60.rename(columns={
        'ประเภทสินเชื่อ': 'cont_group',
        'int.rte': 'interest_flat_rate',
        'วันเริ่มต้น': 'first_due_date',
        'วันสิ้นสุด': 'last_due_date',
        'สถานะสัญญา': 'over_due_period'
    })

    zf50 = pd.read_excel(source_file3, engine="openpyxl")

    zf50 = zf50[["เลขที่สัญญา","วันทำสัญญา","ลูกค้า","เลขที่จดทะเบียน VAT","หมายเลขถัง","เงินต้น",
                "จำนวนงวด","ค่างวด"]].copy()

    for col in ["เลขที่สัญญา", "ลูกค้า", "เลขที่จดทะเบียน VAT"]:
        zf50[col] = zf50[col].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, float) and x.is_integer() else str(x) if pd.notnull(x) else np.nan)
        
    zf50["วันทำสัญญา2"] = zf50["วันทำสัญญา"]

    zf50 = zf50.rename(columns={
        'วันทำสัญญา': 'cont_date',
        'ลูกค้า': 'cust_code',
        'เลขที่จดทะเบียน VAT': 'ID',
        'หมายเลขถัง': 'chassis_no',
        'วันทำสัญญา2': 'actual_date',
        'เงินต้น': 'principal',
        'จำนวนงวด': 'period',
        'ค่างวด': 'installment'
    })

    WM['เลขที่สัญญา'] = WM['เลขที่สัญญา'].apply(lambda x: str(int(x)) if pd.notnull(x) else np.nan)

    zf20_1['interest_accrued'] = pd.to_numeric(zf20_1['interest_accrued'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)

    zf20_1['เลขที่สัญญา'] = zf20_1['เลขที่สัญญา'].apply(lambda x: str(int(x)) if pd.notnull(x) else np.nan)

    zf20_2['เลขที่สัญญา'] = zf20_2['เลขที่สัญญา'].apply(lambda x: str(int(x)) if pd.notnull(x) else np.nan)

    # convert to float

    columns_to_convert = ["installment_over_due","installment_current"]  

    for col in columns_to_convert:
        zf20_2[col] = zf20_2[col].str.replace(",", "", regex=False)  # remove commas
        zf20_2[col] = pd.to_numeric(zf20_2[col], errors='coerce')    # convert to floatzf20_2[columns_to_convert] = zf20_2[columns_to_convert].astype(float)
        
        zf60['เลขที่สัญญา'] = zf60['เลขที่สัญญา'].apply(lambda x: str(int(x)) if pd.notnull(x) else np.nan)
        


    import pandas as pd

    def merge_column_with_same_ID(df1, df2, df3, df4, df5, id_column, how='left'):
        # Normalize column names
        for df in [df1, df2, df3, df4, df5]:
            df.columns = df.columns.str.strip().str.lower()
        id_column = id_column.lower()

        for i, df in enumerate([df1, df2, df3, df4, df5], start=1):
            if id_column not in df.columns:
                raise ValueError(f"DataFrame {i} does not contain '{id_column}'")

        # df2 aggregation: sum interest_accrued, take last next_due_date
        df2_sum_col = 'interest_accrued'
        df2_last_col = 'next_due_date'

        agg_funcs = {}
        if df2_sum_col in df2.columns:
            agg_funcs[df2_sum_col] = 'sum'
        if df2_last_col in df2.columns:
            agg_funcs[df2_last_col] = 'last'

        df2_agg = df2.groupby(id_column, as_index=False).agg(agg_funcs) if agg_funcs else df2.drop_duplicates(subset=[id_column])

        # df3 aggregation: sum all numeric columns except id_column
        df3_numeric_cols = [col for col in df3.columns if col != id_column and pd.api.types.is_numeric_dtype(df3[col])]
        df3_agg = df3.groupby(id_column, as_index=False)[df3_numeric_cols].sum() if df3_numeric_cols else df3.drop_duplicates(subset=[id_column])

        # Deduplicate df4 and df5
        df4_dedup = df4.drop_duplicates(subset=[id_column])
        df5_dedup = df5.drop_duplicates(subset=[id_column])

        # Merge step by step
        merged = df1.merge(df2_agg, on=id_column, how=how)
        merged = merged.merge(df3_agg, on=id_column, how=how)
        merged = merged.merge(df4_dedup, on=id_column, how=how)
        merged = merged.merge(df5_dedup, on=id_column, how=how)

        return merged


    merged_df = merge_column_with_same_ID(WM, zf20_1, zf20_2, zf60, zf50, id_column)

    merged_df['Ref. Aging'] = merged_df['เลขที่สัญญา']


    merged_df['send_bill_status'] =  "Y"

    merged_df['period_installment'] =  "1"

    source_file = merged_df

    # Change this to the actual name
    df_template = pd.read_excel(destination_file, sheet_name=target_sheet)

    # Step 1: Expand df_template to match the number of rows in source_file
    if df_template.empty:
        df_template = pd.DataFrame(columns=df_template.columns)  # keep column names
        df_template = df_template.reindex(index=range(len(source_file)))  # add blank rows
        
    column_mapping = {
        'เลขที่สัญญา': 'cont_no',
        'cont_date': 'cont_date',
        'cust_code': 'cust_code',
        'id': 'ID',
        'chassis_no': 'chassis_no',
        'cont_group': 'cont_group',
        'send_bill_status': 'send_bill_status',##
        'branch_code': 'branch_code',
        'actual_date': 'actual_date',
        'principal': 'principal',
        'interest_flat_rate': 'interest_flat_rate',
        'period_installment': 'period_installment',##
        'period': 'period',
        'installment': 'installment',
        'payment_day': 'payment_day',
        'first_due_date': 'first_due_date',
        'last_due_date': 'last_due_date',
        'current_balance': 'current_balance',
        'interest_accrued': 'interest_accrued',
        'next_due_date': 'next_due_date',
        'installment_current': 'installment_current',
        'installment_interest': 'installment_interest',
        'installment_over_due': 'installment_over_due',
        'dpd': 'dpd',
        'over_due_period': 'over_due_period',
        'Ref. Aging': 'Ref. Aging'
    }


    # Step 2: Copy over mapped data
    for src_col, dest_col in column_mapping.items():
        if src_col in source_file.columns and dest_col in df_template.columns:
            df_template[dest_col] = source_file[src_col].values  # row-wise copy

    # Load workbook
    book = load_workbook(destination_file)

    # Create a Pandas ExcelWriter using the openpyxl engine
    with pd.ExcelWriter(destination_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # No need to assign writer.book anymore
        df_template.to_excel(writer, sheet_name=target_sheet, index=False)
        
    end_time = time.time()
    print(f"Total time: {end_time - start_time:.2f} seconds")