"""
excel处理类
"""


from openpyxl import Workbook
from openpyxl import load_workbook

class ExcelUtil(object):

    def __init__(self,file_name,sheet_name=None,sheet_index=0):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.sheet_index = sheet_index

    def create_excel(self,data_dict):
        """创建excel"""
        workbook = Workbook()
        sheet = None
        if self.sheet_name is None:
            sheet = workbook.create_sheet(index=self.sheet_index)
        else:
            sheet = workbook.create_sheet(title=self.sheet_name,index=self.sheet_index)
        for key in data_dict:
            sheet[key]=data_dict[key]
        workbook.save(self.file_name)
        workbook.close()

    def load_excel(self):
        self.workbook = load_workbook(self.file_name)
        self.sheet = None
        if self.sheet_name is None:
            self.sheet = self.workbook.worksheets[self.sheet_index]
        else:
            self.sheet = self.workbook[self.sheet_name]

    def get_data_by_cell_no(self,cell_no):
        """根据编号获取单元格的值"""
        self.load_excel()
        return self.sheet[cell_no].value

    def get_data_by_row_col_no(self,row_no,col_no):
        """根据行、列索引获取单元格的值"""
        self.load_excel()
        return self.sheet.cell(row_no,col_no).value

    def set_data_by_cell_no(self,cell_no,cell_value):
        """根据编号修改单元格的值"""
        self.load_excel()
        self.sheet[cell_no]=cell_value
        self.workbook.save(self.file_name)

    def set_data_by_row_col_no(self,row_no,col_no,cell_value):
        """根据行、列索引修改单元格的值"""
        self.load_excel()
        self.sheet.cell(row_no,col_no,cell_value)
        self.workbook.save(self.file_name)

    def get_data_by_row_no(self, row):
        """获取整行的所有值，传入行数"""
        self.load_excel()
        columns = self.sheet.max_column
        rowdata = []
        for i in range(1, columns + 1):
            cellvalue = self.sheet.cell(row=row, column=i).value
            rowdata.append(cellvalue)
        return rowdata

    def get_row_no_by_cell_value(self,cell_value,col_no):
        """根据单元格值获取行号"""
        self.load_excel()
        rowns = self.sheet.max_row
        row_no = -1
        for i in range(rowns):
            if self.sheet.cell(i+1,col_no).value == cell_value:
                row_no = i+1
                break
            else:
                continue
        return row_no
