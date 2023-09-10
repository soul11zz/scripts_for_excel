from ExcelBulk import SP1, State, SDC
import sys
import os
import configparser



conf = configparser.ConfigParser()
script_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
conf.read(f"{script_directory}/conf.ini")

BIDTYPE = 1
new_bids = []



# ===================  adders ===================

def add_new_percentage(row:dict, symbol: bool, num: int):
    global new_bids
    bid = float(row["data"]["ag"])
    num = int(num)
    new_bid = 0
    if num == 999:
        new_bid = 0
    else:
        if symbol:
            new_bid = bid + num
        else:
            new_bid = bid - num

    print("Current percentage = ", bid)
    print("Changed percentage = ", new_bid, "Changed on ", num)
    print("==================")
    new_bids.append([int(row["row"]) + 2, row["data"]["ak"], new_bid])
    

def add_new_bid(row: dict, symbol: bool, num: int):
    global new_bids, BIDTYPE
    if BIDTYPE == 1:
        if row["data"]["ab"] == '':
            bid = float(row["data"]["aa"])
        else:
            bid = float(row["data"]["ab"])
    elif BIDTYPE == 2:
        bid = float(row["data"]["z"])
        
    num = int(num)
    new_bid = 0
    if symbol:
        new_bid = bid + ((bid * num) / 100)
    else:
        new_bid = bid - ((bid * num) / 100)

    print("Current Bid = ", bid)
    print("Changed Bid = ", new_bid, "Changed on ", num)
    print("==================")
    new_bids.append([int(row["row"]) + 2, row["data"]["ak"], new_bid])


# ===================  checkers ===================
def check_productTargetingExpression(row):
    if row["data"]["ah"] == "complements": return True
    return False
    
def check_acos(row):
    global confi
    acos = float(row["data"]["ar"]) * 100

    clicks = int(row["data"]["ak"])
    orders = int(row["data"]["ao"])
    
    if acos == 0:
        if check_productTargetingExpression(row):
            return False
        if 30 > clicks > 0:
            return False
        if clicks > 30:
            add_new_bid(row, False, conf["DEFAULT"]["acos0"])
            return True
        
    if 10 >= acos > 0:
        if orders < 2:
            add_new_bid(row, True, conf["DEFAULT"]["acos0_10__orders_0_2"])
            return True
        if orders >= 2:
            add_new_bid(row, True, conf["DEFAULT"]["acos0_10__orders_2"])
            return True
        
    if 17 >= acos > 10:
        add_new_bid(row, True, conf["DEFAULT"]["acos10_17"])
        return True
    
    if 24 >= acos > 17:
        return False

    if 40 > acos > 24:
        if clicks < 30:
            add_new_bid(row, False, conf["DEFAULT"]["acos24_40__clicks_0_30"])
            return True
        if clicks > 30:
            add_new_bid(row, False, conf["DEFAULT"]["acos24_40__clicks_30"])
            return True

    if acos >= 40:
        if clicks < 30:
            add_new_bid(row, False, conf["DEFAULT"]["acos40__clicks_0_30"])
            return True
        if clicks > 30:
            add_new_bid(row, False, conf["DEFAULT"]["acos40__clicks_30"])
            return True

def check_placment_acos(row):
    global confi
    acos = float(row["data"]["ar"]) * 100

    orders = int(row["data"]["ao"])

    if acos <= 17:
        if orders <2:
            return False
        if orders >= 2:
            add_new_percentage(row, True, conf["PERCENTAGE"]["acos0_17__orders_2"])
            return True

    if acos >= 24:
        add_new_percentage(row, False, conf["PERCENTAGE"]["acos_more_24"])
        return True


# ==================== loops ====================
def loop(data: list):
    x = 0
    for row in data:
        if check_acos(row): x+=1

    print(x)


def state_work(data: list):
    print(len(data))
    global new_bids
    state = State("state")
    state.open()
    state = state.get_list_data()
    name = []
    for i in state:
        name.append(i[0])
        
    for row in data:
        cell = row["data"]["v"]
        if cell in name:
            status = row["data"]["r"]
            if str(status) != str(state[name.index(cell)][-1]):
                print(cell, status)
                print(name[name.index(cell)])
                print(state[name.index(cell)])
                print('====================')
                new_bids.append([int(row["row"]) + 2, row["data"]["ak"], str(state[name.index(cell)][-1])])

def state_work2(data: list):
    print(len(data))
    global new_bids
    state = State("state")
    state.open()
    state = state.get_list_data()
    name = []
    for i in state:
        name.append(i[0])
        
    for row in data:
        cell = row["data"]["v"]
        if cell in name:
            status = row["data"]["p"]
            if str(status) != str(state[name.index(cell)][-1]):
                print(cell, status)
                print(name[name.index(cell)])
                print(state[name.index(cell)])
                print('====================')
                new_bids.append([int(row["row"]) + 2, row["data"]["ak"], str(state[name.index(cell)][-1])])


def calculate_procent_change(data: list):
    global new_bids
    for row in data:
        if row["data"]["ab"] == '':
            bid = float(row["data"]["aa"])
        else:
            bid = float(row["data"]["ab"])
        newbid = float(row["data"]["au"])
        d = (newbid - bid) / bid
        new_bids.append([int(row["row"]) + 2, row["data"]["ak"], str(d)])
            
def main(arg: str):
    global new_bids, BIDTYPE
    x = SP1(arg)
    x.open()


    # first ex
    r_data = x.get_dict_data()
    data = x.sorting_first(r_data)
    BIDTYPE = 1
    loop(data)
    x.writelist_row_newbid(new_bids, 1)


    # second ex
    data = x.sorting_second(r_data)
    BIDTYPE = 2
    new_bids = []
    loop(data)
    x.writelist_row_newbid(new_bids, 2)

    # third ex
    data = x.sorting_third(r_data)
    BIDTYPE = 1
    new_bids = []
    loop(data)
    x.writelist_row_newbid(new_bids, 1)
   

    # forty ex
    data = x.sorting_forty(r_data)
    new_bids = []
    calc = 0
    for i in data:
        if check_placment_acos(i): calc+=1
    x.writelist_row_newbid(new_bids, 3)

    # new state for SPC
    data = x.sorting_fifth(r_data)
    new_bids = []
    data = x.sorting_fifth(data)
    state_work(data)
    x.writelist_row_newbid(new_bids, 4)
    x.save_file()

    # calculate Percent change
    x = SP1(arg)
    x.open()
    new_bids = []
    r_data = x.get_dict_data()
    data = x.sorting_six(r_data)
    print("len of bids ", len(data))
    calculate_procent_change(data)
    x.writelist_row_newbid(new_bids, 5)
    
    
    x.save_file()
    
    # new state for SDC
    c = SDC(arg)
    c.open()
    new_bids = []
    c_data = c.get_dict_data()
    data = c.sorting_fifth(c_data)
    state_work2(c_data)
    c.writelist_row_newbid(new_bids, 1)

    c.save_file()

    
if __name__ == "__main__":
    file_name = input("type your file name: ")
    main(file_name)
