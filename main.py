from ExcelClass import Excel, AllInspoExport, GSC
import re


# algo
# if gsc word in "Site descriprion"
# if "Content" has link ->
# 1) link is on isparation: we continuer
# 2) link is on other post: we ignore/ delete

# else: continuer

# next step
# if numbers in the GSC url are the same in this post collections: ignore them
# text should be one to one!!
# if text is one to one:

# save text to Anchor Text
# save Site description in Updated Content with link from GSC url

# we do it only with one post and we continuer with next GSC words

# if "From to" or "Anchor Text" or "Updated Content" not free - ignore


def from_to(data: list, obj):
    pattern = "[0-9]{4}"
    for i in data:
        collections = re.findall(pattern, i["Collections"])
        m = re.findall(pattern, i["Content"])
        if m:
            x = " ".join(collections)
            i["From to"] = f"from {x} to {m[0]}"
            obj.write_cell(int(i["row"]) + 1, 13, i["From to"])


def update_content(all_inspo: list, all_gsc: list, inspo):
    def check_description(row: dict, pattern: str):
        # should fix decor in decorator it shouldn't be find
        if type(row["Site description"]) == str:
            res = re.findall(pattern, row["Site description"])
            if res:
                return res

    def check_url(row: dict):
        if type(row["Content"]) == str:
            first_pattern = "https://www.soulandlane.com/"
            res = re.findall(first_pattern, row["Content"])
            if res:
                pattern = "https://www.soulandlane.com/inspiration/"
                res = re.findall(pattern, row["Content"])
                if res:
                    return True
                else:
                    return False
            else:
                return True

    # query and url
    word = fr"(?i){all_gsc[1]['query']}"
    word = fr"(?i)behind these"
    url = all_gsc[1]["url"]
    
    for row in all_inspo:
        res = check_description(row, word)
        if res != None:
            print(row["row"], res)
            print(check_url(row))
            


def main():
    name="Halloween Internal Links.xlsx"
    ignore_rows_list = []

    def get_data_from_excel(obj):
        excel = obj(name)
        excel.open_file()
        return excel.get_dict_data(), excel

    print("Main file is run")
    all_inspo, inspo = get_data_from_excel(AllInspoExport)
    all_gsc, _ = get_data_from_excel(GSC)
    update_content(all_inspo, all_gsc, inspo)


    from_to(all_inspo, inspo)              
    inspo.save_file()
    
            
if __name__ == "__main__":
    main()
