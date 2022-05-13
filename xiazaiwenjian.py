import urllib.request  # url request
import re  # regular expression
import os  # dirs
import time

'''
url 下载网址
pattern 正则化的匹配关键词
Directory 下载目录
'''



def BatchDownload(url, pattern, Directory):
    # 拉动请求，模拟成浏览器去访问网站->跳过反爬虫机制
    headers = {'User-Agent',
               'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'}
    opener = urllib.request.build_opener()
    opener.addheaders = [headers]
    print('--------------------------')
    # 获取网页内容
    content = opener.open(url).read().decode('Utf-8')
    print(content)
    # 构造正则表达式，从content中匹配关键词pattern
    raw_hrefs = re.findall(pattern, content, 0)

    # set函数消除重复元素
    hset = set(raw_hrefs)

    # 下载链接
    for href in hset:
        # 之所以if else 是为了区别只有一个链接的特别情况
        if (len(hset) > 1):
            link = url + href[0]
            filename = os.path.join(Directory, href[0])
            print("正在下载", filename)
            urllib.request.urlretrieve(link, filename)
            print("成功下载！")
        else:
            link = url + href
            filename = os.path.join(Directory, href)
            print("正在下载", filename)
            urllib.request.urlretrieve(link, filename)
            print("成功下载！")

        # 无sleep间隔，网站认定这种行为是攻击，反反爬虫
        time.sleep(1)

BatchDownload('http://basic.10jqka.com.cn/api/stock/export.php?export=debt&type=year&code=600600', None, 'E:/财务分析/啊啊啊')