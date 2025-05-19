import pandas as pd

try:
    # 尝试读取Excel文件
    df = pd.read_excel('text_info.xlsx', sheet_name='Unfilled')
    
    # 打印列名
    print("Excel文件'text_info.xlsx'中'Unfilled'表格的列名:")
    for i, col in enumerate(df.columns):
        print(f"{i+1}. {col}")
    
    # 检查特定列是否存在
    important_columns = ['Source', 'note', 'verified', 'error']
    print("\n检查重要列是否存在:")
    for col in important_columns:
        exists = col in df.columns
        print(f"列 '{col}': {'存在' if exists else '不存在'}")
    
    # 打印前几行数据的样例
    print("\n数据预览 (前3行):")
    print(df.head(3))
    
except Exception as e:
    print(f"读取Excel文件时出错: {e}")
    
    # 尝试列出所有工作表
    try:
        xls = pd.ExcelFile('text_info.xlsx')
        print(f"\n文件中的工作表: {xls.sheet_names}")
        
        # 如果没有'Unfilled'工作表，尝试读取第一个工作表
        if 'Unfilled' not in xls.sheet_names and len(xls.sheet_names) > 0:
            first_sheet = xls.sheet_names[0]
            print(f"\n尝试读取第一个工作表 '{first_sheet}':")
            df = pd.read_excel('text_info.xlsx', sheet_name=first_sheet)
            print(f"工作表 '{first_sheet}' 的列名:")
            for i, col in enumerate(df.columns):
                print(f"{i+1}. {col}")
    except Exception as inner_e:
        print(f"尝试列出工作表时出错: {inner_e}") 