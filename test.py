# Example: 创建Excel文件
import pandas as pd
from cicc_excel.excelwriter import ExcelWriter

df = pd.read_excel('horrible_example_v0.2.xlsx')
df2 = pd.read_excel('cicc_excel_test.xlsx')
wb = ExcelWriter('formated_demo_v0.2.xlsx', ch_font="楷体", num_font="Arial", en_font="Times New Roman")
wb.load_data(df)
wb.write_data('Test1')
wb.collapse_col(1,5,'Test1')

wb.save()