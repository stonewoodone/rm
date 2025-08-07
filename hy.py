import pandas as pd
import os

# 指定文件夹路径
folder_path = "无人值守化验月报"

# 创建空列表存储所有数据框
all_dfs = []

# 需要保留的列索引
columns_to_keep = [0, 1, 2, 3, 4, 6, 9, 10, 11, 13]

# 新的列名
new_column_names = [
    '序号', '公司名称', '来煤量', '化验日期', '全水Mt', 
    '灰分空干基Aad', '挥发份Vdaf', '固定碳', '全硫', '发热量'
]

# 遍历文件夹中的所有xls文件
for filename in os.listdir(folder_path):
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        
        try:
            # 读取文件，跳过前两行，使用第三行作为列名
            df = pd.read_excel(file_path, skiprows=2)
            print(f"正在处理文件: {filename}")
            
            # 删除最后一行
            df = df.iloc[:-1]
            
            # 选择指定的列
            df = df.iloc[:, columns_to_keep]
            
            # 重命名列
            df.columns = new_column_names
            
            # 确保公司名称列为字符串类型
            df['公司名称'] = df['公司名称'].astype(str)
            
            # 在处理之前打印数据类型信息
            #print("公司名称列的数据类型：", df['公司名称'].dtypes)
            #print("公司名称列的前几行数据：", df['公司名称'].head())
            
            # 添加供应商列（提取公司名称中括号前的部分）
            df['供应商全称'] = df['公司名称'].str.split('（').str[0]  # 处理中文括号
            df['供应商全称'] = df['供应商全称'].str.split('(').str[0]  # 处理英文括号
            #将化验日期转换为日期型,格式为'yyyy-mm-dd'
            df['化验日期'] = pd.to_datetime(df['化验日期'], format='%Y-%m-%d', errors='coerce')
            #添加报表月份和供应商列，将化验日期和供应商全称合并
            df['报表月份供应商'] = df['化验日期'].dt.strftime('%Y-%m') + '-' + df['供应商全称']
            #df['报表月份供应商'] = df['化验日期'] + '-' + df['供应商全称']
            # 检查数据是否为空
            if df.empty:
                print(f"警告：处理后的数据为空 {filename}")
                continue
                
            # 将处理后的数据框添加到列表中
            all_dfs.append(df)
            
        except Exception as e:
            print(f"处理文件 {filename} 时出错: {str(e)}")

# 合并所有数据框
if all_dfs:
    final_df = pd.concat(all_dfs, ignore_index=True)
    print(f"最终汇总数据行数: {len(final_df)}")
    
    # 重新排列列顺序，将供应商列放在公司名称后面
    column_order = [
        '序号', '公司名称', '报表月份供应商','供应商全称', '来煤量', '化验日期', '全水Mt', 
        '灰分空干基Aad', '挥发份Vdaf', '固定碳', '全硫', '发热量'
    ]
    final_df = final_df[column_order]
    
    # 添加年月字段，用于分组统计
    final_df['统计月份'] = pd.to_datetime(final_df['化验日期']).dt.strftime('%Y-%m')
    
    # 定义需要计算加权平均的指标
    weighted_columns = ['全水Mt', '全硫', '发热量', '挥发份Vdaf']
    
    # 创建月度加权平均统计函数
    def weighted_average(group, value_column):
        weights = group['来煤量']
        values = group[value_column]
        return (values * weights).sum() / weights.sum()
    
    # 计算月度加权平均
    monthly_stats = []
    for month in sorted(final_df['统计月份'].unique()):
        month_data = final_df[final_df['统计月份'] == month]
        stats = {'统计月份': month, '月度来煤量': month_data['来煤量'].sum()}
        
        # 计算每个指标的加权平均
        for col in weighted_columns:
            stats[f'{col}_加权平均'] = weighted_average(month_data, col)
        
        monthly_stats.append(stats)
    
    # 转换为DataFrame
    monthly_stats_df = pd.DataFrame(monthly_stats)
    
    # 计算年度累计加权平均
    total_stats = {
        '统计月份': '年度累计',
        '月度来煤量': final_df['来煤量'].sum()
    }
    
    # 计算年度各指标加权平均
    for col in weighted_columns:
        total_stats[f'{col}_加权平均'] = weighted_average(final_df, col)
    
    # 将年度统计添加到月度统计中
    monthly_stats_df = pd.concat([monthly_stats_df, pd.DataFrame([total_stats])], ignore_index=True)
    
    # 设置列的显示顺序
    stats_columns = ['统计月份', '月度来煤量'] + [f'{col}_加权平均' for col in weighted_columns]
    monthly_stats_df = monthly_stats_df[stats_columns]
    
    # 计算按公司名称的发热量加权平均
    company_weighted_heat = final_df.groupby('公司名称').agg(
        来煤总量=('来煤量', 'sum'),
        加权平均发热量=('发热量', lambda x: weighted_average(final_df.loc[x.index], '发热量'))
    ).reset_index()
    
    # 创建Excel写入器
    writer = pd.ExcelWriter("化验月报汇总.xlsx", engine='xlsxwriter')
    
    # 写入原始数据
    final_df[column_order].to_excel(writer, index=False, sheet_name='原始数据')
    
    # 写入统计数据
    monthly_stats_df.to_excel(writer, index=False, sheet_name='月度统计')
    
    # 写入公司发热量加权平均数据
    company_weighted_heat.to_excel(writer, index=False, sheet_name='公司发热量加权平均')
    
    # 获取workbook和worksheet对象
    workbook = writer.book
    worksheet_stats = writer.sheets['月度统计']
    worksheet_company = writer.sheets['公司发热量加权平均']
    
    # 设置数字格式
    num_format = workbook.add_format({'num_format': '0.00'})
    
    # 为统计数据添加数字格式
    for col in range(2, len(stats_columns)):  # 跳过'统计月份'和'月度来煤量'列
        worksheet_stats.set_column(col, col, None, num_format)
    
    # 为公司发热量加权平均数据添加数字格式
    worksheet_company.set_column(2, 2, None, num_format)  # 加权平均发热量列
    
    # 保存文件
    writer.close()
    print("汇总完成！文件已保存为'化验月报汇总.xlsx'")
    
    # 对同一供应商全称按照月度和年度，对发热量进行加权平均
    # 首先，确保'化验日期'列的格式为'YYYY-MM'，并将其重命名为'报表月份'
    final_df['报表月份'] = pd.to_datetime(final_df['化验日期']).dt.strftime('%Y-%m')
    
    # 创建一个新的Excel文件，添加多个工作表
    writer = pd.ExcelWriter("化验月报汇总分类.xlsx", engine='xlsxwriter')
    
    # 对'报表月份'和'供应商全称'列进行分组，计算每个月份和供应商的加权平均发热量
    def weighted_avg(group, value_col, weight_col='来煤量'):
        return (group[value_col] * group[weight_col]).sum() / group[weight_col].sum()
    
    weighted_average_heat = final_df.groupby(['报表月份', '供应商全称']).agg(
        加权平均发热量=('发热量', lambda x: weighted_avg(final_df.loc[x.index], '发热量'))
    ).reset_index()
    # 将加权平均发热量结果写入新的Excel工作表
    weighted_average_heat.to_excel(writer, sheet_name='加权平均发热量', index=False)
    
    # 对'报表月份'和'供应商全称'列进行分组，计算每个月份和供应商的加权平均全水Mt
    weighted_average_moisture = final_df.groupby(['报表月份', '供应商全称']).agg(
        加权平均全水Mt=('全水Mt', lambda x: weighted_avg(final_df.loc[x.index], '全水Mt'))
    ).reset_index()
    # 将加权平均全水Mt结果写入新的Excel工作表
    weighted_average_moisture.to_excel(writer, sheet_name='加权平均全水Mt', index=False)
    
    # 对'报表月份'和'供应商全称'列进行分组，计算每个月份和供应商的加权平均全硫
    weighted_average_sulfur = final_df.groupby(['报表月份', '供应商全称']).agg(
        加权平均全硫=('全硫', lambda x: weighted_avg(final_df.loc[x.index], '全硫'))
    ).reset_index()
    # 将加权平均全硫结果写入新的Excel工作表
    weighted_average_sulfur.to_excel(writer, sheet_name='加权平均全硫', index=False)
    
    # 对'报表月份'和'供应商全称'列进行分组，计算每个月份和供应商的加权平均挥发份
    weighted_average_volatile = final_df.groupby(['报表月份', '供应商全称']).agg(
        加权平均挥发份=('挥发份Vdaf', lambda x: weighted_avg(final_df.loc[x.index], '挥发份Vdaf'))
    ).reset_index()
    # 将加权平均挥发份结果写入新的Excel工作表
    weighted_average_volatile.to_excel(writer, sheet_name='加权平均挥发份', index=False)
    
    # 对'报表月份'和'供应商全称'列进行分组，计算每个月份和供应商的加权平均灰份
    weighted_average_ash = final_df.groupby(['报表月份', '供应商全称']).agg(
        加权平均灰份=('灰分空干基Aad', lambda x: weighted_avg(final_df.loc[x.index], '灰分空干基Aad'))
    ).reset_index()
    # 将加权平均灰份结果写入新的Excel工作表
    weighted_average_ash.to_excel(writer, sheet_name='加权平均灰份', index=False)
    
    # 保存文件并关闭
    writer.close()
    print("分类汇总文件已保存为: 化验月报汇总分类.xlsx")
    
    # 打开现有的Excel文件
    writer = pd.ExcelWriter("化验月报汇总分类.xlsx", engine='openpyxl', mode='a')
    
    # 对'供应商全称'列进行分组，计算累计加权平均发热量
    cumulative_weighted_average_heat = final_df.groupby('供应商全称').agg(
        累计加权平均发热量=('发热量', lambda x: weighted_avg(final_df.loc[x.index], '发热量'))
    ).reset_index()
    # 将累计加权平均发热量结果写入新的工作表
    cumulative_weighted_average_heat.to_excel(writer, sheet_name='累计加权平均发热量', index=False)
    
    # 对'供应商全称'列进行分组，计算累计加权平均全水Mt
    cumulative_weighted_average_moisture = final_df.groupby('供应商全称').agg(
        累计加权平均全水Mt=('全水Mt', lambda x: weighted_avg(final_df.loc[x.index], '全水Mt'))
    ).reset_index()
    # 将累计加权平均全水Mt结果写入新的工作表
    cumulative_weighted_average_moisture.to_excel(writer, sheet_name='累计加权平均全水Mt', index=False)
    
    # 对'供应商全称'列进行分组，计算累计加权平均全硫
    cumulative_weighted_average_sulfur = final_df.groupby('供应商全称').agg(
        累计加权平均全硫=('全硫', lambda x: weighted_avg(final_df.loc[x.index], '全硫'))
    ).reset_index()
    # 将累计加权平均全硫结果写入新的工作表
    cumulative_weighted_average_sulfur.to_excel(writer, sheet_name='累计加权平均全硫', index=False)
    
    # 对'供应商全称'列进行分组，计算累计加权平均挥发份
    cumulative_weighted_average_volatile = final_df.groupby('供应商全称').agg(
        累计加权平均挥发份=('挥发份Vdaf', lambda x: weighted_avg(final_df.loc[x.index], '挥发份Vdaf'))
    ).reset_index()
    # 将累计加权平均挥发份结果写入新的工作表
    cumulative_weighted_average_volatile.to_excel(writer, sheet_name='累计加权平均挥发份', index=False)
    
    # 保存文件并关闭
    writer.close()
    print("累计加权平均值已添加到: 化验月报汇总分类.xlsx")


else:
    print("未找到可处理的文件！")


