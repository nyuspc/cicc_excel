## Install

`pip install cicc_excel`

## How to Use

```python
from cicc_excel.excelwriter import ExcelWriter

# create new writer
# default font is 宋体, Arial, Times New Roman
# to setup a percentage column, put a % in column name
wb = ExcelWriter('export_file_name.xlsx', ch_font="中文字体", num_font="Number Font", en_font="English Font", first_row_color='#EEECE1', first_row_height=20)

#load data into writer in pandas dataframe
wb.load_data(df)

#set highlight columns by name
#color in #FFFFFF format
wb.set_hl_col_by_names('col_name', 'sheet_name', 'background_color', 'color')
wb.set_hl_col_by_names('col_name2', 'sheet_name', 'background_color2', 'color2')
#write dataframe to xlsx file
wb.write_data('sheet_name')

#hide cols
wb.hide_col(5,12,'sheet_name')

#collapse cols (have a + on top)
wb.collapse_col(5,12,'sheet_name')

#freeze cell
wb.freeze(1,4, 'sheet_name')

#set hightlight columns for another sheet
#default background color is yellow
wb.load_data(df2)
wb.set_hl_col_by_names('col_name', 'another_sheet_name')
wb.write_data('another_sheet_name')

#another sheet
wb.load_data(df3)
wb.write_data('last_sheet_name')

#save file
wb.save()
```