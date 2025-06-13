import pandas as pd
import os
import re


# Load province mapping
base_path = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(base_path, "province_code.xlsx")
Transfer = pd.read_excel(excel_path, usecols="B,C,D")
Transfer.columns = ['province_name', 'province_abbr', 'province_code']


# Clean mapping data
Transfer['province_name'] = Transfer['province_name'].astype(str).str.strip()
Transfer['province_abbr'] = Transfer['province_abbr'].astype(str).str.strip()
def transfer_Province(reg_no):
    if not isinstance(reg_no, str):
        return '-'

    reg_no = reg_no.strip()

    # Remove trailing period
    if reg_no.endswith('.'):
        reg_no = reg_no[:-1]

    # Remove all spaces and dashes (clean it first)
    reg_no_cleaned = re.sub(r'[\s\-]', '', reg_no)

    # Extract the last group of Thai letters (non-digit) from the end
    suffix_match = re.findall(r'[^\d]+$', reg_no_cleaned)
    if not suffix_match:
        return '-'

    suffix = suffix_match[0].strip().replace('.', '')
    suffix = suffix.strip('จังหวัด')
    if suffix == '' or suffix.isdigit():
        return '-'
    if suffix == 'กทม':
        return '10'

    # Check 2-character province abbreviation match first
    if len(suffix) == 2:
        match = Transfer[Transfer['province_abbr'] == suffix]
    else:
        # Otherwise try full name match (e.g., เชียงราย, กรุงเทพมหานคร)
        match = Transfer[Transfer['province_name'].str.contains(suffix)]

    if not match.empty:
        return match.iloc[0]['province_code']
    else:
        return '-'


def prepare(df):
    if 'เลขทะเบียน' in df.columns:
        df.rename(columns={'เลขทะเบียน': 'reg_no'}, inplace=True)
    if 'reg_no' in df.columns:
        df['province_code'] = df['reg_no'].apply(transfer_Province)
    return df

