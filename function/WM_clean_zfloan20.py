import pandas as pd
from Migration_to_Template_9_WL import dowload_df

def WM_clean_zfloan20(rawfile) :
    #"zfloan20_04.2025 ไฟล์ดิบ.txt"
    # Read Excel file - the warning doesnt afect work
    df = dowload_df(rawfile,sheet_index=0)
    # --- Remove duplicate columns and print them for debugging ---
    dupes = df.columns[df.columns.duplicated()].tolist()
    if dupes:
        print("Duplicate columns found and removed:", dupes)
    df = df.loc[:, ~df.columns.duplicated()]
    # --- End duplicate column handling ---
    df['การหักล้าง'] = pd.to_datetime(df['การหักล้าง'], errors='coerce')
    doc_col = 'เลขเอกสาร'
    df[['เลขที่สัญญ', 'เลขเอกสาร', 'Clrng doc.', 'การหักล้าง']].isnull().sum()
    def clean_dataframe(df):
        # Step 1: Drop rows where 'เลขเอกสาร1' is empty (NaN or empty string)
        df = df[df['เลขเอกสาร'].notna() & (df['เลขเอกสาร'] != '')]


        # Step 2: Separate rows where 'Clrng doc.1' is empty (these are kept no matter what)
        doc_empty = df[df['Clrng doc.'].isna() | (df['Clrng doc.'] == '')]

        # Step 3: Rows where 'Clrng doc.1' has a value
        doc_has_value = df[df['Clrng doc.'].notna() & (df['Clrng doc.'] != '')].copy()

        # Step 4: Apply your year/month filter on 'การหักล้าง1':
        # Keep rows where year > 2025 OR (year == 2025 AND month > 4)
        mask = (doc_has_value['การหักล้าง'].dt.year > 2025) | \
            ((doc_has_value['การหักล้าง'].dt.year == 2025) & (doc_has_value['การหักล้าง'].dt.month > (pd.Timestamp.now().month)-1))

        doc_has_value = doc_has_value[mask]

        # Combine both sets of rows
        clean_df = pd.concat([doc_empty, doc_has_value], ignore_index=True)

        return clean_df

    cleaned_data = clean_dataframe(df)
    cleaned_data[cleaned_data['Clrng doc.'].notna()][['เลขที่สัญญ', 'เลขเอกสาร', 'Clrng doc.', 'การหักล้าง']]
    cleaned_data[['เลขที่สัญญ', 'เลขเอกสาร', 'Clrng doc.', 'การหักล้าง']].isnull().sum()

    cleaned_data.to_excel(r"WM_Temp6_cleaned.xlsx", index=False)
    print("Exported to WM_Temp6_cleaned.xlsx")

    # Remove commas and convert to float for '        ดบ.ในงวด'
    col = '        ดบ.ในงวด'
    if col in cleaned_data.columns:
        cleaned_data[col] = cleaned_data[col].astype(str).str.replace(',', '', regex=False)
        cleaned_data[col] = pd.to_numeric(cleaned_data[col], errors='coerce')

    # Sum the column
    sum_column_a = cleaned_data['        ดบ.ในงวด'].sum()

    print("Sum of column ดบ.ในงวด:", sum_column_a)

    # Convert to datetime
    cleaned_data['Pstng Date'] = pd.to_datetime(cleaned_data['Pstng Date'], dayfirst=True)

    # Keep only rows where date <= 30.04.2025
    cleaned_dataV2 = cleaned_data[cleaned_data['Pstng Date'] <= pd.Timestamp.today().replace(day=1) + pd.offsets.MonthEnd(0)]

    print(cleaned_dataV2)

    cleaned_dataV2.to_excel(r"WM_Temp6_cleanedV2.xlsx", index=False)
    print("Exported to WM_Temp6_cleanedV2.xlsx")

    # Remove commas and convert to float for '              เงินต้น'
    col2 = '              เงินต้น'
    if col2 in cleaned_dataV2.columns:
        cleaned_dataV2[col2] = cleaned_dataV2[col2].astype(str).str.replace(',', '', regex=False)
        cleaned_dataV2[col2] = pd.to_numeric(cleaned_dataV2[col2], errors='coerce')

    # Sum the column
    sum_column_a = cleaned_dataV2['              เงินต้น'].sum()

    print("Sum of column ดบ.ในงวด:", sum_column_a)

    sum_column = cleaned_dataV2['        ดบ.ในงวด'].sum()
    print(sum_column)