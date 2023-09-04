from ExcelClass import Excel



class Rules(Excel):    
    def __init__(self):
        super().__init__("settings")
    
    def open_file(self, name_of_wb: str):
        self.wb = openpyxl.load_workbook(f"./{self.fileName}.xlsx")
        self.ws = self.wb[neme_of_wb]
        self._print_rows()

    def return_object(self):
        pass
