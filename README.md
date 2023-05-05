## Install

`pip install cicc_excel`

## How to Use

```python
# create new writer
# default font is 宋体, Arial, Times New Roman
wb = ExcelWriter('export_file_name.xlsx', ch_font="中文字体", num_font="Number Font", en_font="English Font")

#load data into writer in pandas dataframe
wb.load_data(df)

#set highlight columns by name
wb.set_hl_col_by_names(['col_name'], 'sheet_name')

#write dataframe to xlsx file
wb.write_data('sheet_name')

#hide cols
wb.hide_col(5,12,'sheet_name')

#freeze cell
wb.freeze(1,4, 'sheet_name')

#set hightlight columns for another sheet
wb.load_data(df2)
wb.set_hl_col_by_names(['col_name'], 'another_sheet_name')
wb.write_data('another_sheet_name')

#another sheet
wb.load_data(df3)
wb.write_data('last_sheet_name')

#save file
wb.save()
```