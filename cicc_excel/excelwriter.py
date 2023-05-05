"""
CICC WM Excel Writer for exporting pandas dataframe data to excel in CICC format.


@author: Pengcheng Song
@Mail: smth_spc@hotmal.com
"""
import pandas as pd
import re
import xlsxwriter

class ExcelWriter(object):
    """
    Class for writing data to an xlsx file.
    """
    def __init__(self, filename, en_font='Arial', ch_font='宋体', num_font='Arial'):
        """
        Initialize an open workbook and add a worksheets.
        """
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(filename, {'nan_inf_to_errors':True})
        self.styles = {
            'normal':{},
            'hl':{}
        }
        self.hl_cols = {}
        #Setting Styles
        self.header_format = self.workbook.add_format({'border':None,
                                    'bold': True,
                                    'font_size': 10,
                                    'font_name': ch_font,
                                    'bg_color': '#EEECE1',
                                    'align': 'center',
                                    'valign': 'vcenter'})
        self.styles['normal']['number_format'] = self.workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '#,##0'
        })
        self.styles['normal']['worktime_format'] = self.workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '#,##0.00'
        })
        self.styles['normal']['percent_format'] = self.workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '0.00%'
        })
        self.styles['normal']['chinese_format'] = self.workbook.add_format({
            'font_name': ch_font,
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        self.styles['normal']['english_format'] = self.workbook.add_format({
            'font_name': en_font,
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        self.styles['normal']['sn_format'] = self.workbook.add_format({
            'font_name': en_font,
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'bold': True
        })
        self.styles['normal']['date_format'] = self.workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'font_name': en_font,
            'num_format':'yyyy/mm/dd'
        })
        self.styles['normal']['default_format'] = self.workbook.add_format({
            'font_size':10,
            'align': 'left',
            'valign': 'vcenter'
        })
        # Set Highlight Styles
        self.styles['hl']['number_format'] = self.workbook.add_format({
            'font_size': 10,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '#,##0'
        })
        self.styles['hl']['worktime_format'] = self.workbook.add_format({
            'font_size': 10,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '#,##0.00'
        })
        self.styles['hl']['percent_format'] = self.workbook.add_format({
            'font_size': 10,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'align': 'right',
            'valign': 'vcenter',
            'font_name': num_font,
            'num_format': '0.00%'
        })
        self.styles['hl']['chinese_format'] = self.workbook.add_format({
            'font_name': ch_font,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter'
        })
        self.styles['hl']['english_format'] = self.workbook.add_format({
            'font_name': en_font,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'font_size': 10,
            'valign': 'vcenter',
            'align': 'left'
        })
        self.styles['hl']['sn_format'] = self.workbook.add_format({
            'font_name': en_font,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'bold': True
        })
        self.styles['hl']['date_format'] = self.workbook.add_format({
            'font_size': 10,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'align': 'right',
            'valign': 'vcenter',
            'font_name': en_font,
            'num_format':'yyyy/mm/dd'
        })
        self.styles['hl']['default_format'] = self.workbook.add_format({
            'font_size':10,
            'bg_color': '#EEECE1',
            'font_color': '#9B3519',
            'align': 'left',
            'valign': 'vcenter'
        })
        #set default width
        self.column_width = {
            'number': 15,
            'worktime': 10,
            'date': 10,
            'percent': 8,
            'text': 15,
            'sn': 8,
            'default': 10
        }
    
    def load_data(self, data):
        """
        load data into workbook
        """
        self.set_global_format()
        self.data = data
    
    def write_data(self, sheet_name='Sheet1'):
        """
        write data into workbook
        """
        ws = self.workbook.add_worksheet(sheet_name)
        headers = list(self.data.columns)
        # add header
        for col_num, header_title in enumerate(headers):
            ws.write(0, col_num, header_title)
        self.set_first_row_format(sheet_name)
        date_columns = []
        num_columns = []
        text_columns = []
        percent_columns = []
        worktime_columns = []
        sn_columns = []
        for col_name in self.data.columns:
            if pd.api.types.is_datetime64_dtype(self.data[col_name]):
                date_columns.append(col_name)
            elif '%' in col_name:
                percent_columns.append(col_name)
            elif '工号' in col_name or '编号' in col_name or '编码' in col_name:
                sn_columns.append(col_name)
            elif 'yr' in col_name or '工时' in col_name:
                worktime_columns.append(col_name)
            elif pd.api.types.is_numeric_dtype(self.data[col_name]):
                num_columns.append(col_name)
            else:
                text_columns.append(col_name)

        # add data
        for col_num, column_name in enumerate(self.data.columns):
            #set high light column
            style = self.styles['normal']
            if sheet_name in self.hl_cols and column_name in self.hl_cols[sheet_name]:
                style = self.styles['hl']
            #set width
            if column_name in date_columns:
                ws.set_column(col_num, col_num, self.column_width['date'], None)
            elif column_name in sn_columns:
                ws.set_column(col_num, col_num, self.column_width['sn'], None)
            elif column_name in num_columns:
                ws.set_column(col_num, col_num, self.column_width['number'], None)
            elif column_name in percent_columns:
                ws.set_column(col_num, col_num, self.column_width['percent'], None)
            elif column_name in worktime_columns:
                ws.set_column(col_num, col_num, self.column_width['worktime'], None)
            elif column_name in text_columns:
                ws.set_column(col_num, col_num, self.column_width['text'], None)
            else:
                ws.set_column(col_num, col_num, self.column_width['default'], None)
            #write cells    
            for row_num, cell_value in enumerate(self.data[column_name]):
                if pd.isnull(cell_value):
                    ws.write(row_num + 1, col_num, '')
                elif column_name in date_columns:
                    ws.write(row_num + 1, col_num, cell_value, style['date_format'])
                elif column_name in sn_columns:
                    ws.write(row_num + 1, col_num, cell_value, style['sn_format'])
                elif column_name in num_columns:
                    ws.write(row_num + 1, col_num, cell_value, style['number_format'])
                elif column_name in percent_columns:
                    ws.write(row_num + 1, col_num, cell_value, style['percent_format'])
                elif column_name in worktime_columns:
                    ws.write(row_num + 1, col_num, cell_value, style['worktime_format'])
                elif column_name in text_columns:
                    if re.search('[\u4e00-\u9fa5]+', str(cell_value)):
                        ws.write(row_num + 1, col_num, cell_value, style['chinese_format'])
                    else:
                        ws.write(row_num + 1, col_num, cell_value, style['english_format'])
                else:
                    ws.write(row_num + 1, col_num, cell_value, style['default_format'])

    def add_worksheet(self, name):
        """
        Add a new worksheets to the workbook.
        """
        self.workbook.add_worksheets(name)

    def save(self):
        """
        Save workbook to the filename given at __init__.
        """
        self.workbook.close()

    def close(self):
        """
        Close the workbook.
        """
        self.workbook.close()
    
    def add_format(self, format):
        """
        Add a new format to the workbook.
        """
        return self.workbook.add_format(format)
    
    def set_first_row_format(self, sheet_name='Sheet1'):
        """
        Add a format to the first row.
        """
        ws = self.workbook.get_worksheet_by_name(sheet_name)
        if ws is not None:
            ws.set_row(0, 20, self.header_format)
            ws.autofilter(0, 0, ws.dim_rowmax,ws.dim_colmax)
        else:
            print("Error: worksheets", sheet_name, "not found")

    def set_global_format(self):
        """
        Set the global format for all worksheets.
        """
        df_format = self.workbook.formats[0]
        df_format.set_font_size(10)
    
    def freeze(self, row=1, col=1, sheet_name='Sheet1'):
        """
        Freeze the first row and first column.
        """
        ws = self.workbook.get_worksheet_by_name(sheet_name)
        if ws is not None:
            ws.freeze_panes(row, col)
        else:
            print("Error: worksheets", sheet_name, "not found")
    
    def autofit(self, sheet_name='Sheet1'):
        """
        Fit the width of all columns.
        """
        ws = self.workbook.get_worksheet_by_name(sheet_name)
        if ws is not None:
            ws.autofit()
        else:
            print("Error: worksheets", sheet_name, "not found")

    def hide_col(self, s_col=0, end_col=0, sheet_name='Sheet1'):
        """
        Hide acolumns.
        """
        ws = self.workbook.get_worksheet_by_name(sheet_name)
        if ws is not None:
            ws.set_column(s_col, end_col, None, None, {'hidden': True})
        else:
            print("Error: worksheets", sheet_name, "not found")

    def collapse_col(self, s_col=0, end_col=0, sheet_name='Sheet1'):
        """
        Hide acolumns.
        """
        ws = self.workbook.get_worksheet_by_name(sheet_name)
        if ws is not None:
            ws.set_column(s_col, end_col, None, None, {'collapsed': True})
        else:
            print("Error: worksheets", sheet_name, "not found")
    
    def set_hl_col_by_names(self, col_names, sheet_name):
        """
        Set col by col name
        """
        self.hl_cols[sheet_name] = col_names



