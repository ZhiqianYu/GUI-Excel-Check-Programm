#Code from yu-zhiqian@outlook.com, ZhiqianYu@github

import pandas as pd

def update_vodafone_rechnung(file_path):
    # 读取 Excel 文件
    xls = pd.ExcelFile(file_path)
    df1 = xls.parse('Aktuell12.02.2025')
    df2 = xls.parse('Vodafone Rechnung')

    # 检查重复
    if df1.duplicated(subset=['Mobile number']).any():
        print("Warning: Duplicate mobile numbers found. Keeping only the first occurrence.")
        df1 = df1.drop_duplicates(subset=['Mobile number'], keep='first')

    # 创建匹配字典
    lookup_dict = df1.set_index('Mobile number').to_dict(orient='index')

    # 遍历 Vodafone Rechnung 表
    df2['Department Update'] = 'not found'
    df2['Name Update'] = 'not found'

    for index, row in df2.iterrows():
        mobile_number = str(row['Mobile number'])  # 确保手机号是字符串
        match = next((key for key in lookup_dict if mobile_number in str(key)), None)

        if match:
            df2.at[index, 'Department Update'] = lookup_dict[match]['Department']
            df2.at[index, 'Name Update'] = lookup_dict[match]['Employee Name']

    # 保存 Excel
    output_path = file_path.replace('.xlsx', 'Mobile_Liste_updated.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Aktuell12.02.2025', index=False)
        df2.to_excel(writer, sheet_name='Vodafone Rechnung', index=False)

    print(f"Updated file saved as {output_path}")

    # 调用函数
file_path = "Mobile Liste.xlsx"  # 请替换为实际文件路径
update_vodafone_rechnung(file_path)