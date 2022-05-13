'''
行业分类
方法：query_stock_industry

入参
code：A股股票代码，sh或sz.+6位数字代码，或者指数代码，如：sh.601398。sh：上海；sz：深圳。可以为空；
date：查询日期，格式XXXX-XX-XX，为空时默认最新日期。

返回参数
参数名称	参数描述	算法说明
updateDate	更新日期
code	证券代码
code_name	证券名称
industry	所属行业
industryClassification	所属行业类别

'''

import baostock as bs
import pandas as pd

# 登陆系统
lg = bs.login(user_id="anonymous", password="123456")
# 显示登陆返回信息
print('login respond error_code:'+lg.error_code)
print('login respond  error_msg:'+lg.error_msg)

# 获取行业分类数据
rs = bs.query_stock_industry()
# rs = bs.query_stock_basic(code_name="浦发银行")
print('query_stock_industry error_code:'+rs.error_code)
print('query_stock_industry respond  error_msg:'+rs.error_msg)
# 打印输出
industry_list = []
while (rs.error_code == '0') & rs.next():
    # 获取一条记录，将记录合并在一起
    industry_list.append(rs.get_row_data())
result = pd.DataFrame(industry_list, columns=rs.fields)
# 结果集输出到csv文件
# result_profit.to_csv("D:\\profit_data.csv", encoding="gbk", index=False)

# 登出系统
bs.logout()