# -*- coding: utf-8 -*-
# @Author: monkey-hjy
# @Date:   2021-02-24 17:18:02
# @Last Modified by:   monkey-hjy
# @Last Modified time: 2021-02-24 17:23:17
import os
from datetime import datetime, time
import xlwt

import requests
from GetToken import GetToken
import random
import pandas as pd
import win32com.client as win32



# 随机UA头
USER_AGENT = [
    "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
    "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
    "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
    "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
    "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
    "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
    "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
    "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
    "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
    "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
    "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
    "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52",
]
#配置位置
#输出地址
outputdir = os.getcwd()
#最少粉丝数
min_fans_num = 1000
execl = xlwt.Workbook(encoding='utf-8',style_compression=0)
#key_words 列表
key_words=['Decentralized Exchange min_faves:10 -filter:links','tradfi min_faves:10 -filter:links', 'synthetix min_faves:10 -filter:links', 'synthetic assets min_faves:10 -filter:links', 
        '(blockchain OR synthetic OR uniswap OR crypto OR cryptocurrency) -RT -retweet -follow -retweets -follows -following -followers min_faves:10 -filter:links']

KOLS=[]
need_token = 1
ii=0
profiles=''
def exporttToExcl(tweets,users,fileName):

        sheet = execl.add_sheet(fileName[0:16]+str(need_token),cell_overwrite_ok=True)
        i=1
        global profiles
        sheet.write(0,0,'推文ID') 
        sheet.write(0,1,'发文时间') 
        sheet.write(0,2,'内容') 
        sheet.write(0,3,'用户名') 
        sheet.write(0,4,'账号') 
        sheet.write(0,5,'粉丝数') 
        sheet.write(0,6,'profile') 
        sheet.write(0,7,'点赞数') 
        sheet.write(0,8,'回复数') 
        sheet.write(0,9,'Investors') 
        sheet.write(0,10,'Funder') 
        sheet.write(0,11,'Twitter handle') 
        sheet.write(0,12,'Twitter link') 
        sheet.write(0,13,'# of followers') 
        # sheet.write(0,14,'Cause') 
        sheet.write(0,14,'Location') 
        # sheet.write(0,16,'Types') 
        
        key_ids=[]

        for key in tweets: 
            user_id = tweets.get(key).get('user_id_str')
            if user_id in key_ids:
                continue
            key_ids.append(user_id)
            cuser=users.get(user_id)
            ctweet=tweets.get(key)
            profile=cuser.get('description')
            
            #最小粉丝数
            if cuser.get('followers_count')<min_fans_num:
                continue
            #排除已有的kol
            if cuser.get('screen_name') in KOLS:
                continue
           
           
            sheet.write(i,0,key) 
            sheet.write(i,1,tweets.get(key).get('created_at')) 
            sheet.write(i,2,tweets.get(key).get('full_text')) 
            sheet.write(i,3,cuser.get('name')) 
            sheet.write(i,4,cuser.get('screen_name')) 
            sheet.write(i,5,cuser.get('followers_count')) 
            sheet.write(i,6,profile) 
            profilelink=profile+cuser.get('screen_name')
            if 'investor' in profilelink or ('Investor' in profilelink) or ('INVESTOR'in profilelink):
                sheet.write(i,9,'Y') 
            if 'funder' in profilelink or 'Funder' in profilelink or 'FUNDER'in profilelink:
                sheet.write(i,10,'Y') 
            KOLS.append(cuser.get('screen_name'))
            profiles+= ' '+ ctweet.get('full_text')
            sheet.write(i,7,ctweet.get('favorite_count')) 
            sheet.write(i,8,ctweet.get('retweet_count')) 
            sheet.write(i,11,cuser.get('screen_name')) 
            sheet.write(i,12,'https://twitter.com/'+cuser.get('screen_name')) 
            sheet.write(i,13,cuser.get('followers_count')) 

            sheet.write(i,14,cuser.get('location')) 
            i+=1
def exportPostToExcl(tweets,users,fileName):
        username=fileName[6:len(fileName)-1]
        # sheets=execl.sheet_names();
        # if(username in sheets):
        #     return
        
        sheet = execl.add_sheet(username,cell_overwrite_ok=True)
        i=1
        global profiles
        sheet.write(0,0,'推文ID') 
        sheet.write(0,1,'发文时间') 
        sheet.write(0,2,'内容') 
        # sheet.write(0,3,'用户名') 
        sheet.write(0,4,'账号') 
        # sheet.write(0,5,'粉丝数') 
        # sheet.write(0,6,'profile') 
        sheet.write(0,7,'点赞数') 
        sheet.write(0,8,'回复数') 
        
        key_ids=[]

        for key in tweets: 
            user_id = tweets.get(key).get('user_id_str')
            # if user_id in key_ids:
            #     continue
            key_ids.append(user_id)
            cuser=users.get(user_id)
            ctweet=tweets.get(key)
            profile=cuser.get('description')
            #sha
            if cuser.get('screen_name')!=username:
                continue
            # if ctweet.get('favorite_count')<10:
            #     continue
            sheet.write(i,0,key) 
            sheet.write(i,1,tweets.get(key).get('created_at')) 
            sheet.write(i,2,tweets.get(key).get('full_text')) 
            # sheet.write(i,3,cuser.get('name')) 
            sheet.write(i,4,cuser.get('screen_name')) 
            profilelink=profile+cuser.get('screen_name')
            sheet.write(i,7,ctweet.get('favorite_count')) 
            sheet.write(i,8,ctweet.get('retweet_count')) 
            i+=1
        # execl.save(outputdir+'/'+'关键词结果'+'.xls')

class SearchTweet(GetToken):
    """
    根据关键词搜索推文或者用户
    使用游客token进行抓取数据，没有次数限制
    但是需要境外ip。。。（此处没用用到ip，具体详见get_token）
    """
    
    def __init__(self):
        super().__init__()
        self.start = datetime.now()
        # 定义请求头。需要按照下面的代码去获取游客token
        self.headers = {
            'authorization': 'Bearer AAAAAAAAAAAAAAAAAAAAANRILgAAAAAAnNwIzUejRCOuH5E6I8xnZz4puTs%3D1Zv7ttfk8LF81IUq16cHjhLTvJu4FA33AGWWjCpTnA',
            'user-agent': random.choice(USER_AGENT),
            # 'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36',
            # 'x-guest-token': '1418373851956715520',
            'x-guest-token': self.get_token(proxies_ip='104.18.6.138:443'),
        # 'x-csrf-token': '03b42d6ffea1ddc5aa800b75712e6f18b291152fe1de8b895bfcfa7090c4d4f843e09f11c2478675495796a81b20f31326c4bee988faa7fa450dba9227fb9e984834e69f3ace98ffdd1d3dba29daee84'
        }
        # 获取数据的接口1417777806684180480
        self.url = 'https://twitter.com/i/api/2/search/adaptive.json'

    def start_requests(self, search_key, kind,search_type='tweet'):
        """
        开始搜索
        :param search_key: 搜素关键词
        :param search_type: 搜索类别。tweet/推文。   account/用户
        :return:
        """
        #以关键字进行请求时 不田间时间筛选
        if kind=='key':
            params = {
                "q": search_key,
                "count": 300,
                "tweet_search_mode": 'live',
                "tweet_mode": "extended",
        }
        else:
            params = {
                "q": search_key+' until:2021-07-25 since:2020-01-01',
                "count": 1000,
                "tweet_search_mode": 'live',
                "tweet_mode": "extended",
            }
        global need_token
        if search_type == 'account':
            params['result_filter'] = 'user'
        response = requests.get(url=self.url, headers=self.headers, params=params, timeout=100)
        if response.status_code != 200:
            # need_token+=1
            print('返回结果coed'+ str(response.status_code))

            return f"{search_key} ERR  ===  {response}"
        tweets = response.json().get('globalObjects').get('tweets')
        users = response.json().get('globalObjects').get('users')
        if not len(tweets) and not len(users):
            # need_token+=1
            print('未抓到数据'+str(len(tweets))+":"+str(len(users)))
            KOLS.remove(search_key[ 6:len(search_key)-1])
            return f'{search_key}未抓到数据'
        # p = PrettyTable()
        if search_type == 'tweet':
            tweet_id = []
            create_time = []
            full_text = []
            user_name = []
            screen_name = []

            
            for key in tweets:
                tweet_id.append(key)
                create_time.append(tweets.get(key).get('created_at'))
                full_text.append(tweets.get(key).get('full_text'))
                user_id = tweets.get(key).get('user_id_str')
                user_name.append(users.get(user_id).get('name'))
                screen_name.append(users.get(user_id).get('screen_name'))

            # p.add_column(fieldname='内容', column=full_text)
            # p.add_column(fieldname='用户名', column=user_name)
            # p.add_column(fieldname='账号', column=screen_name)
            if kind=='key':
                key_words.remove(search_key)
                exporttToExcl(tweets,users,search_key)
            else:
                exportPostToExcl(tweets,users,search_key)
                print(search_key[ 6:len(search_key)-1]+str(len(KOLS)))
                KOLS.remove(search_key[ 6:len(search_key)-1])

        else:
            user_name = []
            screen_name = []
            description = []
            for key in users:
                user_name.append(users.get(key).get('name'))
                screen_name.append(users.get(key).get('screen_name'))
                description.append(users.get(key).get('description'))
            # p.add_column(fieldname='用户名', column=user_name)
            # p.add_column(fieldname='账号', column=screen_name)
            # p.add_column(fieldname='简介', column=description)
        #此处为打印到控制台的测试数据，可以考虑删除
        return 'OK'



    def run(self,key):
        global need_token
        global ii
        search_key = key_words
        # search_key = ['Decentralized Exchange min_faves:10 -filter:links']
        if(key=='key'):
            for skey in search_key:
                #此处可改成account
                result = self.start_requests(search_key=skey,kind=key, search_type='tweet')
                print(need_token)
            # seprate(profiles)
            # need_token-=1
        else:
            # ii=0
            
            for kol in KOLS:
                #此处可改成account
                result = self.start_requests(search_key='(from:'+kol+')',kind=key, search_type='tweet')
                ii+=1
                if(ii%300==0):
                    break
                print("已经爬取到第"+str(ii)+"个KOL的post "+kol)
            # need_token-=1
        

    def __del__(self):
        end = datetime.now()
        print(f'开始：{self.start}，结束：{end}\n用时：{end-self.start}')


if __name__ == '__main__':
    t = SearchTweet()
    while len(key_words)>0:
        t.run('key')
    execl.save(outputdir+'/'+'关键词得到的KOL列表0'+'.xls')
    execl=xlwt.Workbook(encoding='utf-8',style_compression=0)
    while len(KOLS)>0:
        print(len(KOLS))
        t.run('kol')
        # time.sleep(10)
        
        # execl=xlwt.Workbook()
    execl.save(outputdir+'/'+'KOL所有的post结果'+'all1'+'.xls')
     
    
    #转化 为xlsx可以被xlrd读取
    fname = os.getcwd()+ '\关键词得到的KOL列表0.xls'
    excels = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excels.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excels.Application.Quit()
    
    
