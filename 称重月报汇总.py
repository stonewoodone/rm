import pandas as pd
import os

# 获取所有xls文件
folder_path = "无人值守称重月报"
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xls')]

# 创建一个空的DataFrame列表
dfs = []

# 读取每个xls文件并添加到列表中
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    try:
        # 首先读取Excel文件以获取总行数
        temp_df = pd.read_excel(file_path)
        total_rows = len(temp_df)
        
        # 重新读取文件，跳过前两行，不读取最后一行
        df = pd.read_excel(file_path, 
                          skiprows=2,           # 跳过前两行
                          nrows=total_rows-3)   # 总行数减去前两行和最后一行
        
        # 添加月份信息列
        df['报表月份'] = file.replace('.xls', '')
        
        # 重命名"到厂重量（t）"列为"重量"
        if '到厂重量（t）' in df.columns:
            df = df.rename(columns={'到厂重量（t）': '重量'})
            
        # 添加供应商列（提取供应单位中括号前的部分）
        df['供应商全称'] = df['供应单位'].apply(lambda x: 
            str(x).split('（')[0].split('(')[0].strip() if pd.notnull(x) else '')
        #添加报表月份和供应商列
        df['报表月份供应商'] = df['报表月份'] + '-' + df['供应商全称']
        
        dfs.append(df)
        print(f"成功处理文件: {file}")
    except Exception as e:
        print(f"处理文件 {file} 时出错: {str(e)}")

# 合并所有DataFrame
if dfs:
    combined_df = pd.concat(dfs, ignore_index=True)
    
    try:
        # 选择并重排列顺序，加入供应商列
        selected_columns = ['序号', '报表月份', '供应单位', '供应商全称','报表月份供应商', '运输单位', '车数', '重量']
        combined_df = combined_df[selected_columns]
        
        # 保存合并后的文件
        output_file = "称重月报汇总.xlsx"
        combined_df.to_excel(output_file, index=False)
        print(f"文件已合并并保存为: {output_file}")
    except KeyError as e:
        print(f"错误：找不到列 {e}")
        print("请检查列名是否正确，当前可用的列名有：", combined_df.columns.tolist())
else:
    print("没有找到可处理的Excel文件")
