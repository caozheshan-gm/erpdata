import os
import sys
import pandas as pd
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

    # Filter ERP数据 to only include rows with '料号' present in 实盘数
    df2 = df2[df2['料号'].isin(df1['料号'])]

    # Add a column to indicate the source of the data
    df1['来源'] = '实盘数'
    df2['来源'] = 'ERP数据'

    # Select relevant columns and add any missing columns to ensure consistency
    df1 = df1[['料号', '批号', '客户编号', '客户名称', '品名', '规格', '实盘单位', '实盘', '来源']]
    df2 = df2[['料号', '旧批号', '库位', '库存管理特征', '库存单位', '库存数量', '来源']]

    # Combine data from both sources
    df_combined = pd.concat([df1, df2], axis=0, ignore_index=True)

    # Sort by '料号' and ensure '来源' is sorted with '实盘数' above 'ERP数据'
    df_combined['来源'] = pd.Categorical(df_combined['来源'], categories=['实盘数', 'ERP数据'], ordered=True)
    df_combined.sort_values(by=['料号', '来源'], inplace=True)

    # Calculate '相差数' for each 料号
    df_combined['相差数'] = None
    for 料号, group in df_combined.groupby('料号'):
        实盘总和 = group.loc[group['来源'] == '实盘数', '实盘'].sum()
        库存总和 = group.loc[group['来源'] == 'ERP数据', '库存数量'].sum()
        if not group[group['来源'] == 'ERP数据'].empty:
            df_combined.loc[group.index[0], '相差数'] = 库存总和 - 实盘总和

    # Add a new column '杂发数量' and calculate it
    df_combined['杂发数量'] = None
    for 料号, group in df_combined.groupby('料号'):
        实盘总和 = group.loc[group['来源'] == '实盘数', '实盘'].sum()
        for idx, row in group.iterrows():
            if row['来源'] == 'ERP数据':
                if 实盘总和 > row['库存数量']:
                    df_combined.at[idx, '杂发数量'] = row['库存数量']
                    实盘总和 -= row['库存数量']
                else:
                    df_combined.at[idx, '杂发数量'] = 实盘总和
                    实盘总和 = 0
                    break

    # ['料号', '批号', '客户编号', '客户名称', '品名', '规格', '实盘单位', '实盘', '来源', '旧批号', '库位','库存管理特征', '库存单位', '库存数量', '相差数', '杂发数量']
    za = df_combined[['料号', '批号', '客户名称', '实盘单位', '实盘', '来源', '旧批号', '库位',
       '库存管理特征', '库存单位', '库存数量', '相差数', '杂发数量']]
    # Remove rows where '杂发数量' is NaN and '来源' is 'ERP数据'
    za = za[~((df_combined['来源'] == 'ERP数据') & (df_combined['杂发数量'].isna()))]
    
    料号_counts = za['料号'].value_counts()
    condition1 = pd.isna(za['杂发数量'])
    condition2 = za['料号'].map(料号_counts) > 1
    za.loc[condition1, '库位'] = 'CK08'
    za.loc[condition1, '杂发数量'] = -za.loc[condition1, '实盘']

    za['理由码'] = None
    za['检验否'] = None
    za.loc[:, '理由码'] = '107'
    za.loc[condition2, '检验否'] = 'N'

    za = za.drop(columns=['库存数量','实盘','相差数'])
    za.rename(columns={'杂发数量': '库存数量'}, inplace=True)

    df_combined.to_excel('转换数据调整.xlsx', sheet_name='数据调整', index=False)
    za.to_excel('转换杂发单.xlsx', sheet_name='杂发单', index=False)
    print(f"Process completed successfully. File 转换数据调整.xlsx and 转换杂发单.xlsx has been generated.")

except FileNotFoundError as e:
    print(f"Error: {e}")
    sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    sys.exit(1)
