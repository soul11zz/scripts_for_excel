from ExcelClass import FT
from FromToList import FromToList
import re
import sys
from main import *



def run(args: str):
    name = args
    global f, from_to_num

    def get_data_from_excel(obj):
        excel = obj(name)
        excel.open_file()
        return excel.get_dict_data(), excel

    print("From to script was started")
    all_inspo, inspo = get_data_from_excel(AllInspoExport)
    from_to(all_inspo, inspo)

    ft = FT(name)
    ft.open_file()
    ft.write_data(f.return_list())

    f.print()

if __name__ == "__main__":
    run(sys.argv[1])
