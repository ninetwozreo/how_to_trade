# import twitter_v2
# import tqdm
from se_and_count import exportENDkolExcl, seprate, set_sheet_len
from typing import Sequence
import xlrd
import os
import tweetsc
import win32com.client as win32

finish_query_list = []
bad_query_list = []
thesheet=[]
KOLS=[]
def get_excel_data(path):
    global thesheet
    workbook = xlrd.open_workbook(path)
    sheets= workbook.sheet_names()
    set_sheet_len(len(sheets))
    
    a=0

    for sheet in sheets:
        a+=1


        if(a<0):
            continue
        real_sheet=workbook.sheet_by_name(sheet)
        cols = real_sheet.col_values(2)
        # print(cols)
        posts=''
        KOLS.extend(cols)
        for kol in cols:
            posts+=kol
        seprate(posts,sheet)
        print(str(a)+":"+sheet)

        # print(str(a)+":"+sheet+":"+thesheet)
        
        # start_request(search_key='(from:'+kol+')', search_type='tweet')
        
    posts=''
    
    

if __name__ == '__main__':

    #转化 为xlsx可以被xlrd读取
    fname = os.getcwd()+ '\KOL所有的post结果all1.xls'
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
     

    get_excel_data(fname+"x")
