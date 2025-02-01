import pandas as pd
import os

# 指定文件夹路径
folder_path = "无人值守化验月报"

# 创建空列表存储所有数据框
all_dfs = []

# 需要保留的列索引
columns_to_keep = [1, 2, 3, 4, 5, 7, 10, 11, 12, 14]

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
    
    # 导出到Excel文件
    final_df.to_excel("化验月报汇总.xlsx", index=False)
    print("汇总完成！文件已保存为'化验月报汇总.xlsx'")
else:
    print("未找到可处理的文件！")
