import pandas as pd
import os 

base_path = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(base_path, "Group.xlsx")
Transfer = pd.read_excel(excel_path,usecols='A,B',dtype=str)
Transfer.columns = ['Group_code', 'Group_name']

#count_group
def start(df):
    print(df)
    df['cont_group'] = df['cont_group'].apply(transfer_Group)
    return df

def transfer_Group(count_group):
#check count_group with Group_name and replace it with Group_code then return it
 match = Transfer[Transfer['Group_name'] == count_group]
 if not match.empty:
        return match.iloc[0]['Group_code']
 return None

