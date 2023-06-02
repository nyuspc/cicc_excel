# Example: 创建Excel文件
import pandas as pd
from cicc_excel.excelwriter import ExcelWriter

df = pd.read_excel('df_test.xlsx')
wb = ExcelWriter('formated_demo_v0.2.xlsx', ch_font="楷体", num_font="Arial", en_font="Times New Roman", font_size=15)
wb.load_data(df)
wb.set_hl_col_by_names('判断', 'Test1')
wb.write_data('Test1')
wb.collapse_col(1,5,'Test1')
wb.collapse_col(7,9,'Test1')

wb.save()