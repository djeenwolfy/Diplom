import openpyxl
import os
class Folder(object):
    def __init__(self):
        self.dir_folder=os.getcwd()

class WorkBook(Folder):
    def __init__(self,file_name):
        self.file_name=file_name
        super().__init__()
        self.wb=openpyxl.load_workbook(filename = f'{self.dir_folder}\\{file_name}')

class Sheet(WorkBook):
    def __init__(self, file_name,sheet_name):
        super().__init__(file_name)
        self.sheet=self.wb[sheet_name]

    def replacement(self,data_location_begin,data_location_end,data_storage_space_begin,data_storage_space_end):
        array=[]
        index=0
        for col in self.sheet[data_location_begin:data_location_end]:
            for i in col:
                array.append(i.value)
        #print(array)
        for col in self.sheet[data_storage_space_begin:data_storage_space_end]:
            for i in col:
                self.sheet[i.coordinate]=array[index]
                print(array[index])
                index+=1
        self.wb.save('ТСЖ_КАРТОФЕЛЬНЫЙ_Для_ДИПЛОМАred.xlsx')
    def save(self):
        self.wb.save('ТСЖ_КАРТОФЕЛЬНЫЙ_Для_ДИПЛОМАred.xlsx')
TSJ=Sheet("ТСЖ_КАРТОФЕЛЬНЫЙ_Для_ДИПЛОМА.xlsx","ИПУ")
#TSJ.replacement('E2','F141','C2','D141')
TSJ.save()

