# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import baostock as bs
import pandas as pd
import xlwings as xw
import tkinter as tk
import datetime as dt
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

'''
估值指标
#### 获取沪深A股估值指标(日频)数据 ####
    # close 当日股价
    # peTTM 动态市盈率
    # psTTM 市销率
    # pcfNcfTTM 市现率
    # pbMRQ 市净率
'''
def run():
    print('---------------------执行开始---------------------------')
    lg = bs.login(user_id="anonymous", password="123456")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False #警告关闭
    app.screen_updating = False #屏幕更新关闭
    filePath = selectFile('标题', '选择文件')

    #获取时间
    now = dt.datetime.now()
    before_10year = now.year - 10
    #十年前
    before_10year_day = f'{before_10year}-{now.month}-{now.day}'
    try:
        wb = app.books.open(filePath)
        df_gupiao = pd.read_excel(filePath, sheet_name='股票清单',
                      usecols=['股票名称', '股票代码', '证券交易所'],
                      dtype=str)
        # print('---------------------股票清单---------------------------')
        # print(df_gupiao)

        for index, gupiao in df_gupiao.iterrows():
            gupiao_code = gupiao['股票代码']
            gupiao_jys_code = gupiao['证券交易所']
            #query_history_k_data
            rs = bs.query_history_k_data_plus(gupiao_jys_code+'.'+gupiao_code,
                                     # "date,code,close,peTTM,pbMRQ,psTTM,pcfNcfTTM",
                                     "date,code,close,peTTM,pbMRQ",
                                     start_date=before_10year_day, end_date=dt.datetime.strftime(now, '%Y-%m-%d'),
                                     frequency="d", adjustflag="3")
            result_list = []
            while (rs.error_code == '0') & rs.next():
                # 获取一条记录，将记录合并在一起
                result_list.append(rs.get_row_data())
            result = pd.DataFrame(result_list, columns=rs.fields)
            result['peTTM'] = result['peTTM'].astype(float)
            result['pbMRQ'] = result['pbMRQ'].astype(float)
            len_value = len(result)
            cur_peTTM = result.loc[len_value - 1, ('peTTM')]
            cur_pbMRQ = result.loc[len_value - 1, ('pbMRQ')]
            df_gupiao.loc[index, ('当前价格')] = result.loc[len_value - 1, ('close')]

            # 当前市盈率值所在百分位数值
            df_gupiao.loc[index, ('动态市盈率')] = cur_peTTM
            result.sort_values(by='peTTM', inplace=True)
            result = result.reset_index(drop=True)  # 索引重置
            cur_peTTM_index = 0
            for result_index, row in result.iterrows():
                if row['peTTM'] == cur_peTTM:
                    cur_peTTM_index = result_index
            df_gupiao.loc[index, ('近十年动态市盈率所在百分位')] = cur_peTTM_index / len_value
            # 30%市盈率百分位数值
            df_gupiao.loc[index, ('近十年30%百分位的动态市盈率')] = result['peTTM'].quantile(0.3)

            # 当前市净率值所在百分位数值
            df_gupiao.loc[index, ('市净率')] = cur_pbMRQ
            result.sort_values(by='pbMRQ', inplace=True)
            result = result.reset_index(drop=True)  # 索引重置
            cur_pbMRQ_index = 0
            for result_index, row in result.iterrows():
                if row['pbMRQ'] == cur_pbMRQ:
                    cur_pbMRQ_index = result_index
            df_gupiao.loc[index, ('近十年市净率所在百分位')] = cur_pbMRQ_index / len_value
            # 30%市净率百分位数值
            df_gupiao.loc[index, ('近十年30%百分位的市净率')] = result['pbMRQ'].quantile(0.3)

        # print('---------------------结果---------------------------')
        # print(df_gupiao)
        # 写文件
        try:
            wb.sheets['辅助表'].range('A1').clear()# 清空数据
            wb.sheets['辅助表'].range('A1').options(index=False).value = df_gupiao# 写数据
            wb.save() # 保存
        finally:
            wb.close()
    finally:
        app.quit()
        bs.logout()

    print('---------------------执行结束---------------------------')
    '''
    rs = bs.query_history_k_data("sh.600036",
                                 # "date,code,close,peTTM,pbMRQ,psTTM,pcfNcfTTM",
                                 "date,code,close,peTTM,pbMRQ",
                                 start_date='2012-05-17', end_date='2022-05-11',
                                 frequency="d", adjustflag="3")
    #### 打印结果集 ####
    result_list = []
    while (rs.error_code == '0') & rs.next():
        # 获取一条记录，将记录合并在一起
        result_list.append(rs.get_row_data())

    result = pd.DataFrame(result_list, columns=rs.fields)
    result['peTTM'] = result['peTTM'].astype(float)
    result['pbMRQ'] = result['pbMRQ'].astype(float)
    #### 结果集输出到csv文件 ####
    # result.to_csv("D:\\peTTM_sh.600000_data.csv", encoding="gbk", index=False)
    result['中位数'] = result['peTTM'].median()
    #quantile
    len_value = len(result)
    cur_peTTM = result.loc[len_value - 1, ('peTTM')]
    print('近十年市盈率的30%位数值为：' + str(result['peTTM'].quantile(0.3)))  # 百分位对应的值
    print('当前市盈率值为：' + str(result.loc[len_value - 1, ('peTTM')]))  # 百分位对应的值
    print('近十年市净率的30%位数值为：' + str(result['pbMRQ'].quantile(0.3)))# 百分位对应的值
    print('当前市净率值为：' + str(result.loc[len_value-1, ('pbMRQ')]))  # 百分位对应的值
    result.sort_values(by='peTTM', inplace=True)
    result = result.reset_index(drop=True)  #索引重置
    print(result)
    cur_peTTM_index = 0
    for index, row in result.iterrows():
        if row['peTTM'] == cur_peTTM:
            cur_peTTM_index = index
    print(cur_peTTM_index)
    print(cur_peTTM_index/len_value)
    '''

def selectFile(title, content):
    messagebox.showinfo(title, content)
    filePath = filedialog.askopenfilename()
    if len(filePath) == 0:
        messagebox.showerror('错误', '没有选择文件')
        ex = Exception('没有选择文件')
        raise ex
    return filePath

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    run()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
