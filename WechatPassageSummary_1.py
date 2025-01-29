# coding=gbk
# -*- coding:uft-8 -*-
# https://www.bilibili.com/video/BV1qM411P72g/?spm_id_from=333.337.search-card.all.click&vd_source=80640f970e995e05be64a1b68ef3b120

import json
import requests
import datetime
from time import sleep
import pandas as pd
import os
import openpyxl


def main():
    # Range depends on the number of pages(e.g. 30 Pages = 0-30)
    for i in range(0, 20):
        print('Page {}>>>'.format(i + 1))
        try:
            i *= 5  # Every page there are 5 articles
            url = 'https://mp.weixin.qq.com/cgi-bin/appmsgpublish?sub=list&search_field=null&begin={}&count=5&query=&fakeid={}&type=101_1&free_publish_type=1&sub_action=list_ex&token={}&lang=zh_CN&f=json&ajax=1'.format(
                i, fakeid, token)
            headers = {
                'authority': 'mp.weixin.qq.com',
                'cookie': cookie,
                'referer': 'https://mp.weixin.qq.com/cgi-bin/appmsg?t=media/appmsg_edit_v2&action=edit&isNew=1&type=77&createType=0&token=654631327&lang=zh_CN&timestamp=1738059718845',
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36'
            }
            # 2025/01/29 title+link+updatetime current path publish_page -> publish_list[index] index=0-4-> publish_info -> appmsgex[0]
            res = requests.get(url=url, headers=headers).json()
            res = res['publish_page']
            res = json.loads(res)

            for k in res['publish_list']:
                k = k['publish_info']
                k = json.loads(k)
                k = k['appmsgex'][0]
                title = k['title']  # Get title
                timeStamp = k['update_time']  # Get updatetime
                dateArray = datetime.datetime.fromtimestamp(timeStamp)  # Convert timeStamp into proper time pattern
                time = dateArray.strftime('%Y-%m-%d %H:%M:%S')  # Year-month-day, hour-minute-second
                href = k['link']  # Get link
                dic = {
                    'Title': title,
                    'Time': time,
                    'Link': href
                }
                # Put dic into a list
                if dic not in result:
                    result.append(dic)
                    print(dic)
        except:
            print('error')
            break
        finally:
            if not res['publish_list']:
                break
            sleep(2)


if __name__ == '__main__':
    start = 1
    # fakeid, token are in payload; cookie is in header
    # update your cookie
    fakeid = ''  # target account id
    token = ''  # Your account id
    cookie = ''
    file = 'Account Summery.xlsx'
    result = []
    main()
    # Create the summary of articles
    df = pd.DataFrame(result)
    df.to_excel('Article Summary.xlsx', index=False)
