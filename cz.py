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
        
        # 对不同供应商全称按月及按年进行分类汇总
        # 首先，确保'报表月份'列的格式为'YYYY-MM'
        combined_df['报表月份'] = pd.to_datetime(combined_df['报表月份']).dt.strftime('%Y-%m')
        
        # 对'报表月份'列进行分组，计算每个月份的供应量
        monthly_supply = combined_df.groupby(['报表月份', '供应商全称'])['重量'].sum().reset_index()
        
        # 对'供应商全称'列进行分组，计算每个供应商的累计供应量
        cumulative_supply = combined_df.groupby('供应商全称')['重量'].sum().reset_index()
        
        # 对'报表月份'列进行分组，计算每个月份的平均供应量
        average_supply = combined_df.groupby(['报表月份', '供应商全称'])['重量'].mean().reset_index()
        
        # 对'报表月份'列进行分组，计算每个月份的最大供应量
        max_supply = combined_df.groupby(['报表月份', '供应商全称'])['重量'].max().reset_index()
        
        # 对'报表月份'列进行分组，计算每个月份的最小供应量
        min_supply = combined_df.groupby(['报表月份', '供应商全称'])['重量'].min().reset_index()
        
        # 创建一个新的Excel文件，添加多个工作表
        writer = pd.ExcelWriter("称重月报汇总分类.xlsx", engine='xlsxwriter')
        
        # 将合并后的数据写入第一个工作表
        combined_df.to_excel(writer, sheet_name='合并数据', index=False)
        
        # 将月度供应量写入第二个工作表
        monthly_supply.to_excel(writer, sheet_name='月度供应量', index=False)
        
        # 将累计年度供应量写入第三个工作表
        cumulative_supply.to_excel(writer, sheet_name='累计年度供应量', index=False)
        
        # 将平均供应量写入第四个工作表
        average_supply.to_excel(writer, sheet_name='平均供应量', index=False)
        
        # 将最大供应量写入第五个工作表
        max_supply.to_excel(writer, sheet_name='最大供应量', index=False)
        
        # 将最小供应量写入第六个工作表
        min_supply.to_excel(writer, sheet_name='最小供应量', index=False)
        
        # 保存文件并关闭
        writer.close()
        print("分类汇总文件已保存为: 称重月报汇总分类.xlsx")
    except KeyError as e:
        print(f"错误：找不到列 {e}")
        print("请检查列名是否正确，当前可用的列名有：", combined_df.columns.tolist())
else:
    print("没有找到可处理的Excel文件")
