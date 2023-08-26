from ExcelClass import Excel, AllInspoExport, GSC
import re
import sys


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
    from_to_num = 0
    for i in data:
        if i["Collections"] == None: continue
        if i["Content"] == None: continue
        collections = re.findall(pattern, i["Collections"])
        m = re.findall(pattern, i["Content"])
        if m:
            x = " ".join(collections)
            i["From to"] = f"from {x} to {m[0]}"
            obj.write_cell(int(i["row"]) + 1, 13, i["From to"])
            from_to_num += 1

    obj.save_file()
    return from_to_num

def update_content(all_inspo: list, all_gsc: list, inspo):
    updated_content = 0
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

    def update_content(word:str, url:str):
        for row in all_inspo:
            res = check_description(row, word)
            if (res != None):
                if (row["Anchor Text"] != None): continue
                if check_url(row):
                    url_groupe_id = re.findall(r"[0-9]{4}", url)[0]
                    row_post_collections = re.findall(r"[0-9]{4}", row["Collections"])                
                    if url_groupe_id not in row_post_collections:
                        print(row["row"], res)
                        print(url_groupe_id)
                        print(row_post_collections)
                        inspo.write_cell(int(row["row"]) + 1, 11, str(res[0]))
                        row["Anchor Text"] = str(res[0])
                        string = f"<a href='{url}'>{str(res[0])}</a>"
                        row["Updated Content"] = row["Site description"].replace(str(res[0]), string)
                        inspo.write_cell(int(row["row"]) + 1, 12, row["Updated Content"])
                        return 1
        return 0
                                                 

    # query and url
    for i in all_gsc:
        print(i["row"])
        word = fr"(?i){i['query']}\W"
        url = i["url"]
        try:
            updated_content += update_content(word, url)
        except: continue
            
    inspo.save_file()
    return updated_content
    



def main(args: str):
### name="Halloween Internal Links.xlsx"
    name = args
    ignore_rows_list = []

    def get_data_from_excel(obj):
        excel = obj(name)
        excel.open_file()
        return excel.get_dict_data(), excel

    print("Main file is run")
    all_inspo, inspo = get_data_from_excel(AllInspoExport)
    all_gsc, _ = get_data_from_excel(GSC)

    from_to_num = from_to(all_inspo, inspo)
    print("from_to : " + str(from_to_num))


    updated_content = update_content(all_inspo, all_gsc, inspo)
    

    
    print("from_to : " + str(from_to_num))
    print("updated content : " + str(updated_content))

    
            
if __name__ == "__main__":
    main(sys.argv[1])