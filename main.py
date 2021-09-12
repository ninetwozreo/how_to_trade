# coding=utf-8

import xlwt
import os
import json
import re
import sys
import requests

from utils.log import NOTICE, log, ERROR, RECORD

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.append(BASE_DIR)

import time
import logging
# from utils.parse_until import *
from config import CELERY_BROKER, CELERY_BACKEND, CRAWL_INTERVAL, NUM_STATIC, COUNT_NUM_STATIC, TRAIN_NUM_HEAD
# from db_access import *

from utils.html_downloader import crawl, crawl_law_post
from bs4 import BeautifulSoup
from celery import Celery

# from multiprocessing import Pool, cpu_count

celery_app = Celery('law_engine', broker=CELERY_BROKER, backend=CELERY_BACKEND)
celery_app.conf.update(CELERY_TASK_RESULT_EXPIRES=3600)
num_static1 = NUM_STATIC;
#输出地址
outputdir = os.getcwd()
#运行方法标志
sign = "pkulaw";
#中央chl或地方lar
# stype = "lar";
stype = "chl";
#关键字
keyword ="老旧小区.xlsx";
#每页条
pagesize =100;
test_html=''

# -------日志-------------
logger = logging.getLogger()  # 不加名称设置root logger
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s: - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S')

# 使用FileHandler输出到文件
fh = logging.FileHandler('log.out')
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)

# 使用StreamHandler输出到屏幕
ch = logging.StreamHandler()
ch.setLevel(logging.WARNING)
ch.setFormatter(formatter)

# 添加两个Handler
logger.addHandler(ch)
logger.addHandler(fh)


# 获取到指定类别的所有信息
def get_useful_data(rc_data_html):
    local_res=[]
    datas=BeautifulSoup(rc_data_html,'html.parser').find_all('div',class_='block')
    for one in datas:
        oneData={}
        a=one.find('a')
        oneData["title"] =a.text
        oneData["href"] =a.attrs.get('href')
        local_res.append(oneData)
    return local_res

# 获取网页信息
def get_all_data(url,test_html):
    
    
    if  test_html != '' :
        print("测试数据已存在")
    else:
        test_html=crawl(url)
        pass
    str1=BeautifulSoup(test_html,'html.parser').findAll('script')[31].text

    json1=json.loads(str1[19:len(str1)-1])['contents']['twoColumnBrowseResultsRenderer']['tabs'][1]['tabRenderer']
    json2=json1['content']['sectionListRenderer']['contents'][0]['itemSectionRenderer']['contents'][0]['gridRenderer']['items']
    json3=json2[0]['gridVideoRenderer']['title']['runs']
    dataMap=[]
    print(json3)
    exportToExcl(json3,"YouTube视频列表")

def exportToExcl(res,fileName):
    execl = xlwt.Workbook()

    sheet = execl.add_sheet("测试表名",cell_overwrite_ok=True)
    i=1
    for one in res: 
        sheet.write(i,0,one['text']) 
    execl.save(outputdir+'/'+fileName+'.xlsx')



if __name__ == '__main__':
    url='https://www.youtube.com/c/ChainlinkOfficial/videos'
    sign = "youtube"
    # if (sign == "pkulaw"):
        # get_pku_law()
    if (sign == "youtube"):
        get_all_data(url,test_html)

