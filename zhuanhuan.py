import os
import sys
import pandas as pd
import numpy as np
import openpyxl
try:
    # Check if the input file exists
    input_file = '数据整理.xlsx'
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"The file '{input_file}' is missing in the current directory.")

    # Read data from the input file
    df1 = pd.read_excel(input_file, sheet_name='实盘数')
    df2 = pd.read_excel(input_file, sheet_name='ERP数据')

    # Convert '料号' column to string
    df1["料号"] = df1["料号"].astype(str)
    df2["料号"] = df2["料号"].astype(str)

    # Rename '批号' to '旧批号'
    df2.rename(columns={'批号': '旧批号'}, inplace=True)

    # Merge data
    df = pd.merge(
        df1[['批号', '客户编号', '客户名称', '料号', '品名', '规格', '实盘单位', '实盘']],
        df2[['料号', '旧批号', '库位', '库存管理特征', '库存单位', '库存数量']],
        on='料号',
        how="left"
    )

    # Group by '料号' and calculate differences
    grouped = df.groupby('料号').agg({
        '库存数量': 'sum',
        '实盘': 'first'
    }).reset_index()

    grouped['相差数'] = np.where(
        (grouped['库存数量'] != 0) & pd.notna(grouped['库存数量']),
        grouped['库存数量'] - grouped['实盘'],
        np.nan
    )

    # Merge the difference back into the main DataFrame
    df = df.merge(grouped[['料号', '相差数']], on='料号', how='left')

    # Initialize '杂发数量' as None
    df['杂发数量'] = None

    # Compute '杂发数量' for each group
    for name, group in df.groupby('料号'):
        remaining_shipment = group['实盘'].iloc[0]
        for i, row in group.iterrows():
            if pd.notna(row['库存数量']) and row['库存数量'] != 0:
                diff = remaining_shipment - row['库存数量']
                if diff > 0:
                    df.loc[i, '杂发数量'] = row['库存数量']
                    remaining_shipment = diff
                else:
                    df.loc[i, '杂发数量'] = remaining_shipment
                    break

    # Filter rows with '杂发数量' not null
    df3 = df[df['杂发数量'].notna()]

    # Combine data for the final output
    za = pd.concat([df1, df3], axis=0).sort_values(by='料号')[
        ['批号', '老订单号', '客户名称', '料号', '单位', '实盘', '备注', '旧批号', '库位', 
         '库存管理特征', '库存数量', '杂发数量']
    ]

    # Count occurrences of '料号'
    料号_counts = za['料号'].value_counts()

    # Conditions for updating '杂发数量' and '库位'
    condition1 = pd.isna(za['杂发数量'])
    condition2 = za['料号'].map(料号_counts) > 1
    mask = condition1 & condition2

    za.loc[mask, '库位'] = 'CK08'
    za.loc[mask, '杂发数量'] = -za.loc[mask, '实盘']

    # Drop unnecessary columns
    za = za.drop(columns=['库存数量', '实盘'])

    # Add and update additional columns
    za['理由码'] = None
    za['检验否'] = None
    za.loc[condition2, '理由码'] = '107'
    za.loc[condition2, '检验否'] = 'N'

    # Save results to Excel
    df.to_excel('转换数据调整.xlsx', sheet_name='数据调整', index=False)
    za.to_excel('转换杂发单.xlsx', sheet_name='杂发单', index=False)

    print("Process completed successfully. Files '转换数据调整.xlsx' and '转换杂发单.xlsx' have been generated.")

except FileNotFoundError as e:
    print(f"Error: {e}")
    sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    sys.exit(1)