# from docx import Document
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# from docx.oxml.ns import qn # 中文格式
# from docx.shared import Pt # 磅数
# from docx.shared import Inches # 图片尺寸
import urllib3
import urllib
from bs4 import BeautifulSoup
from urllib.request import urlopen
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from win32com.client import DispatchEx
import pyautogui

# IE浏览器下载button的截图
img = r'./capture.png'

def IEDownload(url):
    ie = DispatchEx('InternetExplorer.Application')
    ie.Navigate(url)

#     最多尝试查找5次，避免死循环
    times = 0
    while times < 5:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=1, button='left', duration=0.01, interval=0.01)
            break
        times += 1

plt.rcParams['font.sans-serif'] = ['SimHei'] # 解决中文乱码问题
plt.rcParams['axes.unicode_minus'] = False # 解决坐标值为负数时无法正常显示负号的问题

c_name = '青岛啤酒'
stock_code = '600600'
st = 2021
et = 2021

# 同花顺网站下载{链接的固定字段:文件名固定字段}
ref = {'main&type=report': 'main_report.xls',
       'main&type=year': 'main_year.xls',
       'main&type=simple': 'main_simple.xls',
       'debt&type=report': 'debt_report.xls',
       'debt&type=year': 'debt_year.xls',
       'benefit&type=report': 'benefit_report.xls',
       'benefit&type=year': 'benefit_year.xls',
       'benefit&type=simple': 'benefit_simple.xls',
       'cash&type=report': 'cash_report.xls',
       'cash&type=year': 'cash_year.xls',
       'cash&type=simple': 'cash_simple.xls'}

urls = {}
for i in ref.keys():
    # 如果原来下载过就略过以节约时间
    if not os.path.exists(f'{stock_code}_{ref[i]}'):
        urls[f'{stock_code}_{ref[i]}'] = str(f'http://basic.10jqka.com.cn/api/stock/export.php?export={i}&code={stock_code}')

# http://quotes.money.163.com/service/zcfzb_600600.html?type=year
#http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=year&code=600600
# https://money.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/600600/ctrl/all.phtml
IEDownload('http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=year&code=600600')
# html1 = urlopen('https://money.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/600600/ctrl/all.phtml')
headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            'Referer' : 'http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=year&code=600600',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36',
            'Host' : 'basic.10jqka.com.cn'
        }

'''
http://stat.10jqka.com.cn/q?id=f10_cwfx_dcsj&ld=browser&size=1920x1080&nj=1&ref=
http://stockpage.10jqka.com.cn/600600/&url=http://stockpage.10jqka.com.cn/600600/finance/#view&cs=1000x3746&ts=1652360719061
'''
# url = 'http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=year&code=600600'
http = urllib3.PoolManager()
response = http.request('Get', 'http://stockpage.10jqka.com.cn/600600/finance/#view', headers=headers)
print(response.data)
# with open('E:/财务分析/啊啊啊/abc.xml', 'wb') as f:
#     f.write(html.data)
#     print(f"下载成功")


# http.request
# opener = urllib.request.build_opener()
# opener.addheaders = [headers]
# urllib.request.urlretrieve('http://stockpage.10jqka.com.cn/600600/finance/#view', 'E:/财务分析/啊啊啊/abc.html')
# pd.read_excel('https://drive.google.com/uc?export=download&id=16cp23cJxeyUfnBHMp-sNCuFNQxe8cqOV')
# soup1 = BeautifulSoup(html1,features='lxml')
# txt = soup1.text
# print(txt)




# windows = 0
# for filename in urls:
    # 每7次调用一次xmlhttp，胆子大可以把这个值设小点
    # if windows % 7 == 0:
    #     XMLHTTP(filename, urls[filename])
    #     if not os.path.exists(f'./{filename}'):
    #         IEDownload(urls[filename])
    #     windows += 1
    # else:
    #     IEDownload(urls[filename])
    #     windows += 1
        # 每7次关闭IE的所有窗口，释放内存
        # if windows % 7 == 0:
        #     time.sleep(0.05)
        #     QuitIE()
        #     time.sleep(0.05)