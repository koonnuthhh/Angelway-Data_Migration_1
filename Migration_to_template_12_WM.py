import pandas as pd
from openpyxl import load_workbook
output_file = "12-ข้อมูลค่างวดค้างชำระสัญญาเงินกู้(แยกเงินต้นดอกเบี้ย)_result.xlsx"
# Declare variable
def Migration_to_template_12_WM(destination_file) :
    source_file = r"WM_Temp6_cleanedV2.xlsx"
    target_sheet = "ค่างวดค้างชำระสัญญาเงินกู้"
    df_template = pd.read_excel(destination_file, sheet_name=target_sheet)

    source_file = pd.read_excel(source_file, engine="openpyxl")

    column_mapping = {
        'เลขที่สัญญ': 'cont_no',
        'วันครบกำหน': 'due_date',
        '          มูลค่าในงวด': 'installment',
        '              เงินต้น': 'principal',
        'หดบ.ในงวด': 'interest'
    }

    if df_template.empty:
        df_template = pd.DataFrame(columns=df_template.columns)  # keep column names
        df_template = df_template.reindex(index=range(len(source_file)))  # add blank rows

    # Step 2: Copy over mapped data
    for src_col, dest_col in column_mapping.items():
        if src_col in source_file.columns and dest_col in df_template.columns:
            df_template[dest_col] = source_file[src_col].values  # row-wise copy
            # Load workbook
    book = load_workbook(destination_file)

    # Create a Pandas ExcelWriter using the openpyxl engine
    with pd.ExcelWriter(output_file , engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # No need to assign writer.book anymore
        df_template.to_excel(writer, sheet_name=target_sheet, index=False)
