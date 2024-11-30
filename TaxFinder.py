import pandas as pd
import re

#--<1. 주소 형식 통일>--

# Load the 'tax' Excel file
tax_df = pd.read_excel(r'C:\Users\prove\Desktop\PropertyTax_2024\PropertyTaxL.xlsx', sheet_name=1)

# Function to remove leading zeros from address
def remove_leading_zeros(address):
    return re.sub(r'\b0+(\d+)', r'\1', address)

# Apply the function to the '주소_텍스트' column
tax_df['주소_텍스트'] = tax_df['주소_텍스트'].apply(remove_leading_zeros)

#--<2. 자산번호 찾기>--

# Load the other Excel files
report_df = pd.read_excel(r'C:\Users\prove\Desktop\PropertyTax_2024\ref2_PropertyReport_2409.xlsx')
old_df = pd.read_excel(r'C:\Users\prove\Desktop\PropertyTax_2024\참고_23년 재산세 납부현황 및 과세적정성 검토결과(예천지사_김상국).xlsx', sheet_name=4)

# 중복 제거
report_df = report_df.drop_duplicates(subset=['주소_텍스트_자산레포트'])
old_df = old_df.drop_duplicates(subset=['주소_텍스트_old'])

# Ensure columns are of string type
report_df['주소_텍스트_자산레포트'] = report_df['주소_텍스트_자산레포트'].astype(str)
tax_df['주소_손질'] = tax_df['주소_손질'].astype(str)
old_df['자산번호_old'] = old_df['자산번호_old'].astype(str)

# Merge 'tax_df' with 'report_df' based on the matching columns
merged_df = tax_df.merge(report_df[['주소_텍스트_자산레포트', '자산번호']], left_on='주소_손질', right_on='주소_텍스트_자산레포트', how='left')

# B 파일을 A의 결과와 다시 병합
merged_df = old_df.merge(old_df[['주소_텍스트_old', '자산번호_old']], left_on='주소_손질', right_on='주소_텍스트_old', how='left')

merged_df.to_excel('01.자산번호 찾기_수수정.xlsx', index=False)

#--<3. 귀속사업 찾기>-- 상세사업명 필터링해서 사업구분은 그냥 복붙하자. 여기 안나오면 다른 지사 사업...
discount_df = pd.read_excel(r'C:\Users\prove\Desktop\PropertyTax_2024\ref3_PropertyDiscount_2407.xlsx')
discount_df['사업명'] = discount_df['사업명'].astype(str)

# 중복 제거
discount_df = discount_df.drop_duplicates(subset=['농지소재지'])

# Merge with discount_df
merged_df = tax_df.merge(old_df[['주소_텍스트_old', '사업명']], left_on='주소_손질', right_on='주소_텍스트_old', how='left')
merged_df = merged_df.merge(discount_df[['농지소재지', '사업명']], left_on='주소_손질', right_on='농지소재지', how='left')

# Save the updated file
merged_df.to_excel('02.사업명 찾기_수수정.xlsx', index=False)

#끝!111