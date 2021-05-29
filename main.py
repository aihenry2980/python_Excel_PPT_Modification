import math
import sys

import pandas as pd
import os


def findVar(db, nameEmployID):
    for x in db.iloc:
        if str(x[0]) == nameEmployID[1]:
            if x[1].strip() == nameEmployID[0]:
                if str(x["5th"])=="nan":
                    return "No Score"
                else:
                    return str(x["5th"])
    return "Not Found"


def excel_find_score(srcFile, dbFile):
    src = pd.read_excel(srcFile,usecols="A")
    db = pd.read_excel(dbFile)

    df = pd.DataFrame(columns=['Name', 'ID', 'Brand','Score'])
    for x in src.data:
        x = x.strip()
        nameEmployID = x.split("/")
        if  len(nameEmployID)==3:
            score = findVar(db, nameEmployID)
            lst = [nameEmployID[0], nameEmployID[1], nameEmployID[2], score]
            # df.append(lst)
            df.loc[len(df)] = lst
        else:
            lst = [x,"","",""]
            # df.append(lst)
            df.loc[len(df)] = lst
    # print(df)
    df.to_excel(os.path.splitext(srcFile)[0]+'_result.xlsx')
    return

def ppt_pinyin(pptFile):

    return


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # jxMain("src.xlsx", "staff list.xlsx")
    if sys.argv[1].endswith(".xlsx"):
        excel_find_score(sys.argv[1], sys.argv[2])
    elif sys.argv[1].endswith(".pptx"):
        ppt_pinyin(sys.argv[1])