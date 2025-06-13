import pandas as pd
import os
import sys
import function.change_Branchcode7_9.branchcode_change as bc
# from transferGroup import transfer_Group
from openpyxl import load_workbook
from function.ColumnMappingFunction import map_excel_columns 
import function.Function_count_group_Wl7.transferGroup as tg

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
base_path = os.path.dirname(os.path.abspath(__file__))
destination_file = os.path.join(base_path, "Template", "7-ข้อมูลสัญญาเช่าซื้อ.xlsx")

def Migration_to_Template_7_WL(source_file_1,source_sheet_1,source_file_2,source_sheet_2,destination_file,destination_sheet) :

    # First mapping (Source 1 into Destination)
    mapped_df1 = map_excel_columns(
        source_file_1,
        destination_file,
        source_sheet_1,
        destination_sheet,
        column_mapping={   # Source 1 mapping
            'cont_no': 'cont_no',
            'cont_date': 'cont_date',
            'cust_code': 'cust_code',
            'Cust_ID': 'Cust_ID',
            'collateral_type': 'collateral_type',
            'chassis_no': 'chassis_no',
            'send_bill_status': 'send_bill_status',
            'collateral_vat_rate': 'collateral_vat_rate',
            'product_price': 'product_price',
            'net_product_price': 'net_product_price',
            'vat_product_price': 'vat_product_price',
            'down_payment': 'down_payment',
            'principal': 'principal',
            'net_principal': 'net_principal (H)',
            'vat_principal': 'vat_principal',
            'interest_year': 'interest_flat_rate',
            'period': 'period',
            'installment': 'installment',
            'net_installment': 'net_installment',
            'vat_installment': 'vat_installment',
            'installment_amount': 'installment_amount (G)',
            'net_installment_amount': 'net_installment_amount',
            'vat_installment_amount': 'vat_installment_amount (J)',        
            'penalty_method': 'penalty_method',
            'penalty_rate': 'penalty_rate',
            'material_group': 'material_group',
            'deferred_interest': 'deferred_interest (I)',
        }
    )

    # Second mapping (Source 2 into Destination)
    mapped_df2 = map_excel_columns(
        source_file_2,
        destination_file,  # Same template to get headers
        source_sheet_2,
        destination_sheet,
        column_mapping={   # Source 2 mapping
            'PRC_GW': 'standard_price',
            'NET_PRC_GW': 'net_standard_price',
            'VAT_PRC_GW': 'vat_standard_price',
            'period_installment': 'period_installment',
            'first_due_date': 'first_due_date',
            'last_due_date': 'last_due_date',
            'sale_code': 'sale_code',
            'billcollector_code': 'billcollector_code',
        }
    )

    # ⚠️ Now MERGE the two mapped DataFrames:
    # Where mapped_df2 is not empty, fill mapped_df1 (destination)
    final_df = mapped_df1.copy()

    for column in mapped_df2.columns:
        if column in final_df.columns:
            # Only replace if mapped_df2 has non-NaN (not empty) values
            final_df[column] = mapped_df2[column].combine_first(final_df[column])

    # Fill the column in the destination DataFrame
    if "checker_code" in final_df.columns:
        final_df["checker_code"] = "4261"
    else:
        final_df["checker_code"] = "4261"


    # Fill the column in the destination DataFrame
    if "penalty_late" in final_df.columns:
        final_df["penalty_late"] = "7"
    else:
        final_df["penalty_late"] = "7"



    # ✅ Export the final merged dataframe
    df = pd.read_excel(source_file_1,usecols='G,I',dtype=str)
    df.columns =['cont_group','branch_code']
    ######logic import file final_df= logic
    df = tg.start(df)
    df = bc.start(df)
    final_df['cont_group'] = df['cont_group']
    final_df['branch_code'] = df['branch_code']

    print(df.columns)
    print(final_df.columns)
    final_df['net_installment_amount_issue']=final_df['net_installment']*final_df['period']
    final_df['vat_installment_amount_issue']=final_df['vat_installment']*final_df['period']
    final_df['เช็ค Diff กับช่อง AC']=final_df['deferred_interest (I)'] + final_df['vat_installment_amount (J)']+final_df['net_principal (H)']

    final_df['deferred_interest_issue']=final_df['net_installment_amount_issue'] - final_df['net_principal (H)']
    final_df['เช็ค AC-AQ'] = final_df['installment_amount (G)'] - final_df['เช็ค Diff กับช่อง AC']
    final_df['เช็ค Diff กับช่อง AC.1'] = final_df['net_principal (H)']+final_df['vat_installment_amount_issue']+final_df['deferred_interest_issue']
    final_df['เช็ค AC-AS'] = final_df['installment_amount (G)']-final_df['เช็ค Diff กับช่อง AC.1']


    final_df.to_excel(
        destination_file,
        index=False,
        sheet_name="ข้อมูลสัญญาเชื้อ"
    )

