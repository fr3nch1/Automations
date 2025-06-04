# Data Separation Script

import pandas as pd 
import re

# load xlsx
file_path = r"___.xlsx"  # update with your actual file path

df = pd.read_excel(file_path, sheet_name="Data")



def extract_estimates(value):
    if pd.isna(value):
        return (None, None, None)
    match = re.match(r'(?P<mid>[\d <]+)\s*\[(?P<low>[\d <]+)\s*-\s*(?P<high>[\d <]+)\]', str(value))
    if match:
        def parse_num(s): return int(re.sub(r'[^\d]', '', s)) if s.strip() else None
        mid = parse_num(match.group('mid'))
        low = parse_num(match.group('low'))
        high = parse_num(match.group('high'))
        return (mid, low, high)
    else:
        return (None, None, None)


mid_df = pd.DataFrame()
low_df = pd.DataFrame()
high_df = pd.DataFrame()
mid_df['Country'] = df['Country']
low_df['Country'] = df['Country']
high_df['Country'] = df['Country']


years = [col for col in df.columns if col != 'Country']
for year in years:
    estimates = df[year].apply(extract_estimates)
    mid_df[year] = [e[0] for e in estimates]
    low_df[year] = [e[1] for e in estimates]
    high_df[year] = [e[2] for e in estimates]

# Save all to a new Excel file with 3 sheets
output_path = "HIV_Estimates_Separated_Deaths.xlsx"
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    mid_df.to_excel(writer, sheet_name='MidEst', index=False)
    low_df.to_excel(writer, sheet_name='LowEst', index=False)
    high_df.to_excel(writer, sheet_name='HighEst', index=False)

print(f"Saved to: {output_path}")
