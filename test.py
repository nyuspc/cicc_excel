# Example: 创建Excel文件
import pandas as pd
from cicc_excel.excelwriter import ExcelWriter

df = pd.read_excel('horrible_example_v0.2.xlsx')
df2 = pd.read_excel('horrible_example_v0.2.xlsx')
wb = ExcelWriter('formated_demo_v0.2.xlsx', ch_font="楷体", num_font="Arial", en_font="Times New Roman")
wb.load_data(df)
wb.set_hl_col_by_names('姓名', 'Test1', '#FFFF00')
wb.write_data('Test1')
wb.hide_col(1,2,'Test1')
wb.freeze(1,4, 'Test1')

wb.set_hl_col_by_names('姓名', 'Test2', '#000000', '#00FF00')
wb.set_hl_col_by_names('总客户数', 'Test2')
wb.load_data(df2)
wb.write_data('Test2')

wb.write_data('Test3')
wb.save()