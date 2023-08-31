from ExcelClass import Excel
import sys
import openpyxl

class SP1(Excel):
    # AQ = %
    # AR = %
    # AL = %
    
    
    def open(self):
        self.wb = openpyxl.load_workbook(f"./{self.fileName}.xlsx")
        self.ws = self.wb['Sponsored Products Campaigns']
        self._print_rows()

    def get_dict_data(self):
        data = self._iter_rows()
        res = []
        x = 0
        headers = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'aa', 'ab', 'ac', 'ad', 'ae', 'af', 'ag', 'ah', 'ai', 'aj', 'ak', 'al', 'am', 'an', 'ao', 'ap', 'aq', 'ar', 'as', 'at']
        for i in data[1:]:
            data_dict = {}
            pos = 0
            for j in i:
                data_dict[headers[pos]] = j
                pos +=1

            res.append({
                "row": x,
                "data" : data_dict
                })
            x+=1
        return self._sorting_first(res)

    def _sorting_first(self, data: dict):
        B = "Product Targeting"
        N = ["01. Auto - Bulk Products", "02. Auto - Single Products", "07. SP Product Targeting"]
        R = "enabled"
        S = "enabled"
        T = "enabled"
        
        return_data = []
        a, b, c = 0, 0, 0
        for row in data:
            if self._check_coll(row, 'b', B):
                a += 1
                if (self._check_coll(row, 'n', N[0])) or (self._check_coll(row, 'n', N[1])) or (self._check_coll(row, 'n', N[-1])):
                    b += 1
                    if self._check_coll(row, 'r', R) and self._check_coll(row, 's', S) and self._check_coll(row, 't', T):
                        c += 1
                        return_data.append(row)
        return return_data

    def _check_coll(self, dictionary: dict, coll: str, coll_data: str):
        if dictionary["data"][coll] == coll_data: return True
        return False

def main(arg: str):
    x = SP1(arg)
    x.open()

    print(len(x.get_dict_data()))

if __name__ == "__main__":
    main(sys.argv[1])
