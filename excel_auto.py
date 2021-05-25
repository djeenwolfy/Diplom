import openpyxl
import os
class Folder(object):
    def __init__(self):
        self.dir_folder=os.getcwd()

class WorkBook(Folder):
    wb=0
    name="T"
    def __init__(self,Name):
        self.name=Name
        super().__init__()
        self.wb=openpyxl.load_workbook(filename = f'{self.dir_folder}\\{Name}')
TSJ=WorkBook("ТСЖ_КАРТОФЕЛЬНЫЙ_Для_ДИПЛОМА.xlsx")

print(TSJ.wb.get_sheet_names())


