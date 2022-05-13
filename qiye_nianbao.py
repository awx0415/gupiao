# from docx import Document
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# from docx.oxml.ns import qn # 中文格式
# from docx.shared import Pt # 磅数
# from docx.shared import Inches # 图片尺寸
from bs4 import BeautifulSoup
from urllib.request import urlopen
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

plt.rcParams['font.sans-serif'] = ['SimHei'] # 解决中文乱码问题
plt.rcParams['axes.unicode_minus'] = False # 解决坐标值为负数时无法正常显示负号的问题

#财务指标函数定义
def hs1(cs1,cs2):
    return round(cs1/cs2,2)

def hs2(cs1,cs2,cs3):
    return round((cs1-cs2)/cs3,2)

def hs3(cs1,cs2,cs3):
    return round((cs1+cs2)/cs3,2)

def hs4(cs1,cs2,cs3):
    return round(cs1/(cs2-cs3),2)

def hs5(cs1,cs2,cs3):
    return round(cs1/(cs2+cs3),2)

def hs6(cs1,cs2,cs3,cs4):
    return round((cs1-cs2-cs3)/cs4,2)

def hs7(cs1,cs2):
    return round((cs1-cs2)/cs1,2)

def change(x,y):
    if x.loc[y,'%s年' % et] < x.loc[y,'%s年' % st]:
        return ('%s由%s年的%.2f万元，变化为%s年的%.2f万元，变化数为%.2f万元,变化率为%.2f' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y,'%s年' % et],round(x.loc[y,'%s年' % et]-x.loc[y,'%s年' % st],2),
        round((x.loc[y,'%s年' % et]-x.loc[y,'%s年' % st])/x.loc[y,'%s年' % st],2)))
    else:
        return ('%s由%s年的%.2f万元，变化为%s年的%.2f万元，变化数为%.2f万元,变化率为%.2f' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y,'%s年' % et],round(x.loc[y,'%s年' % et] - x.loc[y,'%s年' % st],2),
        round((x.loc[y,'%s年' % et] - x.loc[y, '%s年' % st])/x.loc[y, '%s年' % st],2)))

def up(x,y):
    if x.loc[y,'%s年' % et] < x.loc[y,'%s年' % st]:
        return ('%s指标由%s年的%.2f，变化为%s年的%.2f，变化数为%.2f，由此看出该公司的%s指标变弱' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y,'%s年' % et],round(x.loc[y,'%s年' % et]-x.loc[y,'%s年' % st],2),y))
    else:
        return ('%s指标由%s年的%.2f，变化为%s年的%.2f，变化数为%.2f，由此看出该公司的%s指标变强' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y, '%s年' % et],round(x.loc[y,'%s年' % et]-x.loc[y,'%s年'% st],2),y))

def down(x,y):
    if x.loc[y,'%s年' % et] < x.loc[y,'%s年' % st]:
        return ('%s指标由%s年的%.2f，变化为%s年的%.2f，变化数为%.2f，由此看出该公司的%s指标变强' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y,'%s年' % et],round(x.loc[y,'%s年' % et]-x.loc[y,'%s年' % st],2),y))
    else:
        return ('%s指标由%s年的%.2f，变化为%s年的%.2f，变化数为%.2f，由此看出该公司的%s指标变弱' %
        (y,st,x.loc[y,'%s年' % st],et,x.loc[y, '%s年' % et],round(x.loc[y,'%s年' % et]-x.loc[y,'%s年'% st],2),y))

#计算平均资产总额
def pj_data(xm):
    pj_data = []
    for i in range(len(xm)-1):
        if i < len(xm):
            pjz = round((xm[i] + xm[i+1])/2,2)
            pj_data.append(pjz)
    return pj_data

#计算增长率
def zzl_data(xm):
    zzl_data = []
    for i in range(len(xm)-1):
        if i < len(xm):
            zzl = round((xm[i]-xm[i+1])/xm[i+1],2)
            zzl_data.append(zzl)
    return zzl_data

#画图函数
def fig(x,y):
    fig = plt.figure()
    for i in range(len(x.index)):
        plt.plot(x.columns.values,x.iloc[i],label =x.index[i])
        plt.title(y)
        plt.legend(loc ='upper right')
        fig.savefig(r'E:\财务分析\成果展示\%s.png'% y)
    return fig

#生成表格函数
# def data_table(x,y):
#     t = d1.add_table(rows=y.shape[0] + 1, cols=y.shape[1] + 1, style='Table Grid')
#     for n in range(len(x)):
#         t.cell(0, n).text = x[n]
#         for m in range(len(list(y.index))):
#             t.cell(m +1, 0).text = list(y.index)[m]
#         for j in range(y.shape[0]):
#             for i in range(y.shape[1]):
#                 t.cell(j +1, i + 1).text = str(y.iloc[j, i])
#     return t

# c_name = input('请输入公司名称：')
c_name = '青岛啤酒'
# stock_code = input("请输入股票代码：")
stock_code = '600600'
# st = input('请输入开始年份：')
# et = input('请输入结束年份：')
st = 2021
et = 2021
data_list =['zcfzb','lrb','xjllb'] #资产负债表、利润表、现金流量表
adrees ='http://quotes.money.163.com'
y_list =['报告日期']+[str(m)+'年' for m in list(range(int(st)-1,int(et)+1))][::-1]

# 1、获取数据并写入到xlsx中
for i in data_list:
    url ='http://quotes.money.163.com/f10/'+i+'_'+stock_code+'.html?type=year'
    html = urlopen(url)
    soup = BeautifulSoup(html,'lxml')
    div =soup.findAll('div',{'class':'inner_box'})
    df = BeautifulSoup(str(div[0]),features='lxml')
    a = df.findAll('a')
    for each in a:
        if each.string=='下载数据':
            new_html = adrees + each.get("href")
            print(new_html)
            html1 = urlopen(new_html)
    soup1 = BeautifulSoup(html1,features='lxml')
    txt = soup1.text.replace('(万元)','').replace('--','0')
    #写临时文件csv
    csv = open(r'E:\财务分析\数据采集\%s%s.csv' % (stock_code, i), 'w', encoding='utf-8').write(
        txt)

    #读取临时文件csv,并转化成df
    data = pd.read_csv(r'E:\财务分析\数据采集\%s%s.csv' % (stock_code, i))
    list1 = list(range(int(data.columns[-2][:4]), int(data.columns[1][:4]) + 1))[::-1]
    new_list = ['报告日期'] + [str(i) + '年' for i in list1] + [data.columns[-1][:4]]
    data.columns = new_list
    #x写到xlsx文件中
    writer = pd.ExcelWriter(r'E:\财务分析\数据采集\%s%s.xlsx' % (stock_code, i))
    data.to_excel(writer, index=False)
    writer.save()
    writer.close()

# 2.删除csv文件
p_list = os.listdir(r'E:\财务分析\数据采集')
for filename in p_list:
    if filename.endswith('.csv'):
        os.remove(r'E:\财务分析\数据采集\%s' % (filename))

p_list = os.listdir(r'E:\财务分析\数据采集')

# 3.处理资产负债表
fp = r'E:\财务分析\数据采集\%s%s.xlsx' % (stock_code,'zcfzb')
data_zcfzb = pd.read_excel(fp, sheet_name='Sheet1', usecols=y_list)
data_zcfzb = data_zcfzb.fillna(0)
data_zcfzb['报告日期'] = data_zcfzb['报告日期'].str.strip() #删除两边的空格
data_zcfzb = data_zcfzb.set_index('报告日期')
#资产结构主要数据
zc = data_zcfzb.loc[['货币资金','应收账款','预付款项','其他应收款','存货','流动资产合计','长期股权投资','固定资产','在建工程','无形资产',
            '长期待摊费用','非流动资产合计']]

zc2 = data_zcfzb.loc[['货币资金','应收账款','预付款项','其他应收款','存货','长期股权投资','固定资产','在建工程','无形资产',
            '长期待摊费用']]
zc_bh = round(abs(zc2['%s年'%et]-zc2['%s年'%st]),2)
zc1 = round(data_zcfzb.loc[['货币资金','应收账款','预付款项','其他应收款','存货','长期股权投资','固定资产','在建工程','无形资产',
            '长期待摊费用']]/data_zcfzb.loc['资产总计'],2)

#负债结构主要数据
fz = data_zcfzb.loc[['短期借款','应付账款','预收账款','应付职工薪酬','应交税费','其他应付款','流动负债合计','长期借款','长期应付款',
            '非流动负债合计']]
fz2 = data_zcfzb.loc[['短期借款','应付账款','预收账款','应付职工薪酬','应交税费','其他应付款','长期借款','长期应付款']]
fz_bh = round(abs(fz2['%s年'%et]-fz2['%s年'%st]),2)
fz1 = round(data_zcfzb.loc[['短期借款','应付账款','预收账款','应付职工薪酬','应交税费','其他应付款','长期借款','长期应付款']]/data_zcfzb.loc['负债合计'],2)

#股本结构主要数据
qy = data_zcfzb.loc[['实收资本(或股本)','资本公积','盈余公积','未分配利润','所有者权益(或股东权益)合计']]
qy2 = data_zcfzb.loc[['实收资本(或股本)','资本公积','盈余公积','未分配利润']]
qy_bh = round(abs(qy2['%s年'%et]-qy2['%s年'%st]),2)
qy1 = round(data_zcfzb.loc[['实收资本(或股本)','资本公积','盈余公积','未分配利润']]/data_zcfzb.loc['所有者权益(或股东权益)合计'],2)

# 3.处理利润表
fp = r'E:\财务分析\数据采集\%s%s.xlsx' % (stock_code,'lrb')
data_lrb = pd.read_excel(fp, sheet_name='Sheet1', usecols=y_list)
data_lrb = data_lrb.fillna(0)
data_lrb['报告日期'] = data_lrb['报告日期'].str.strip() #删除两边的空格
data_lrb = data_lrb.set_index('报告日期')
#利润表主要数据
lr = data_lrb.loc[['营业总收入','营业收入','其他业务收入','营业总成本','营业成本','其他业务成本','销售费用','管理费用','财务费用','其他业务利润',
            '营业利润','利润总额','所得税费用','净利润']]

# 3.处理现金流量表
fp = r'E:\财务分析\数据采集\%s%s.xlsx' % (stock_code,'xjllb')
data_xjllb = pd.read_excel(fp, sheet_name='Sheet1', usecols=y_list)
data_xjllb = data_xjllb.fillna(0)
data_xjllb['报告日期'] = data_xjllb['报告日期'].str.strip() #删除两边的空格
data_xjllb = data_xjllb.set_index('报告日期')
#现金流量表主要数据
xj = data_xjllb.loc[['经营活动现金流入小计','经营活动现金流出小计','经营活动产生的现金流量净额','投资活动现金流入小计','投资活动现金流出小计',
            '投资活动产生的现金流量净额','筹资活动现金流入小计','筹资活动现金流出小计','筹资活动产生的现金流量净额','现金及现金等价物净增加额']]

#计算偿债能力指标
ldbl=hs1(cs1=data_zcfzb.loc['流动资产合计'],cs2=data_zcfzb.loc['流动负债合计'])                #流动比率
sdbl=hs2(cs1=data_zcfzb.loc['流动资产合计'],cs2=data_zcfzb.loc['存货'],cs3=data_zcfzb.loc['流动负债合计'])          #速动比率
zcfzl=hs1(cs1=data_zcfzb.loc['负债合计'],cs2=data_zcfzb.loc['资产总计'])                                   #资产负债率
gdqybl=hs1(cs1=data_zcfzb.loc['所有者权益(或股东权益)合计'],cs2=data_zcfzb.loc['资产总计'])                            #股东权益比率
cqfzbl=hs1(cs1=data_zcfzb.loc['非流动负债合计'],cs2=data_zcfzb.loc['资产总计'])                             #长期负债比率
cqzwyu=hs4(cs1=data_zcfzb.loc['非流动负债合计'],cs2=data_zcfzb.loc['流动资产合计'],cs3=data_zcfzb.loc['流动负债合计'])   #长期债务与营运资金比率
fzsy=hs1(cs1=data_zcfzb.loc['负债合计'],cs2=data_zcfzb.loc['所有者权益(或股东权益)合计'])                                   #负债与所有者权益比率
czcz=hs5(cs1=data_zcfzb.loc['非流动负债合计'],cs2=data_zcfzb.loc['所有者权益(或股东权益)合计'],cs3=data_zcfzb.loc['非流动负债合计'])   #长期资产与长期资金比率
zbhl=hs1(cs1=data_zcfzb.loc['非流动负债合计'],cs2=data_zcfzb.loc['所有者权益(或股东权益)合计'])                               #资本化比率
zbgdh=hs2(cs1=data_zcfzb.loc['资产总计'],cs2=data_zcfzb.loc['非流动负债合计'],cs3=data_zcfzb.loc['所有者权益(或股东权益)合计'])         #资本固定化比率
cqbl=hs1(cs1=data_zcfzb.loc['负债合计'],cs2=data_zcfzb.loc['所有者权益(或股东权益)合计'])                                     #产权比率
cz=pd.DataFrame({'流动比率':ldbl,'速动比率':sdbl,'资产负债率':zcfzl,'股东权益比率':gdqybl,'长期负债比率':cqfzbl,
                '长期债务与营运资金比率':cqzwyu,'负债与所有者权益比率':fzsy,
                '长期资产与长期资金比率':czcz,'资本化率':zbhl,'资本固定化比率':zbgdh,'产权比率':cqbl})
cz = cz.T
cz_bh = cz.std(axis =1)

# 计算盈利能力指标
zzclr = hs1(cs1 = data_lrb.loc['利润总额'],cs2 = pj_data(xm= data_zcfzb.loc['资产总计']))
zzcjlr = hs1(cs1 = data_lrb.loc['净利润'],cs2 =pj_data(xm= data_zcfzb.loc['资产总计']))
yylr =hs1(cs1 = data_lrb.loc['营业利润'],cs2 =data_lrb.loc['营业总收入'] )
jzcsy = hs1(cs1 = data_lrb.loc['净利润'],cs2 =data_zcfzb.loc['所有者权益(或股东权益)合计'] )
gbbc = hs1(cs1 = data_lrb.loc['净利润'],cs2 =data_zcfzb.loc['实收资本(或股本)'] )
xsml = hs7(cs1=data_lrb.loc['营业收入'],cs2=data_lrb.loc['营业成本'])
yl = pd.DataFrame({'总资产利润率':zzclr,'总资产净利润率':zzcjlr,'营业利润率':yylr,'净资产收益率':jzcsy,'股本报酬率':gbbc,'销售毛利率':xsml})
yl = yl.T
yl_bh = yl.std(axis =1)

#计算运营能力指标
yszkzzl = hs1(cs1=data_lrb.loc['营业收入'],cs2=pj_data(xm=data_zcfzb.loc['应收账款']))   # 应收账款周转率
yszkzzt = round(360/yszkzzl,2)    # 应收账款周转天数
chzzl = hs1(cs1=data_lrb.loc['营业成本'],cs2=pj_data(xm=data_zcfzb.loc['存货']))   # 存货周转率
chzzt = round(360/chzzl,2)    # 存货周转天数
zzzzzl = hs1(cs1=data_lrb.loc['营业总收入'],cs2=pj_data(xm=data_zcfzb.loc['资产总计']))    # 总资产周转率
zzzzzt = round(360/zzzzzl,2)    # 总资产周转天数
ldzczzl = hs1(cs1=data_lrb.loc['营业收入'],cs2=pj_data(xm=data_zcfzb.loc['流动资产合计']))   # 流动资产周转率
ldzczzt = round(360/ldzczzl,2)       # 流动资产周转天数

yy = pd.DataFrame({'应收账款周转率':yszkzzl,'应收账款周转天数':yszkzzt,'存货周转率':chzzl,'存货周转天数':chzzt,
                '总资产周转率':zzzzzl,'总资产周转天数':zzzzzt,'流动资产周转率':ldzczzl,'流动资产周转天数':ldzczzt})
yy = yy.T
yy_bh = yy.std(axis =1)

# 计算成长能力指标
zyyw = zzl_data(xm=data_lrb.loc['营业收入'])  # 主营业务收入增长率
jlrzz = zzl_data(xm=data_lrb.loc['净利润'])  # 净利润增长率
jzzzz = zzl_data(xm=data_zcfzb.loc['所有者权益(或股东权益)合计']) # 净资产增长率
zzzzz = zzl_data(xm=data_zcfzb.loc['资产总计']) # 总资产增长率
czn = pd.DataFrame({'主营业务收入增长率':zyyw,'净利润增长率':jlrzz,'净资产增长率':jzzzz,'总资产增长率':zzzzz})
czn = czn.T
czn.columns = [str(m)+'年' for m in list(range(int(st),int(et)+1))][::-1]
czn_bh = czn.std(axis =1)
writer = pd.ExcelWriter(r'E:\财务分析\成果展示\%s%s年至%s年财务分析基数数据表.xlsx' %(c_name,st,et))

zy = pd.concat([data_zcfzb, data_lrb, data_xjllb]).loc[['流动资产合计','非流动资产合计','资产总计','流动负债合计','非流动负债合计','负债合计','所有者权益(或股东权益)合计','营业总收入',
                '营业总成本','利润总额','净利润']]
zy.to_excel(writer,sheet_name='财报主要数据表')
zc.to_excel(writer,sheet_name='资产结构表')
fz.to_excel(writer,sheet_name='负债结构表')
qy.to_excel(writer,sheet_name='股本结构表')
lr.to_excel(writer,sheet_name='利润表主要数据表')
xj.to_excel(writer,sheet_name='现金流量表主要数据表')
cz.to_excel(writer,sheet_name='偿债能力主要数据表')
yl.to_excel(writer,sheet_name='盈利能力主要数据表')
yy.to_excel(writer,sheet_name='运营能力主要数据表')
czn.to_excel(writer,sheet_name='成长能力主要数据表')
writer.save()
writer.close()

#制作各种所需图
#财务主要数据变化趋势图
fig(x=zy,y='财务主要数据变化趋势图')

#资产结构变化趋势图
fig(x=zc,y='资产结构变化趋势图')

#负债结构变化趋势图
fig(x=fz,y='负债结构变化趋势图')

#股本结构变化趋势图
fig(x=qy,y='股本结构变化趋势图')

#利润表主要数据变化趋势图
fig(x=lr,y='利润表主要数据变化趋势图')

#现金流量表主要数据变化趋势图
fig(x=xj,y='现金流量表主要数据变化趋势图')

#偿债能力主要指标变化趋势图
fig(x=cz,y='偿债能力主要指标变化趋势图')

#盈利主要指标变化趋势图
fig(x=yl,y='盈利能力主要指标变化趋势图')

#运营能力主要指标变化趋势图
fig(x=yy,y='运营能力主要指标变化趋势图')

#成长能力主要指标变化趋势图
fig(x=czn,y='成长能力主要指标变化趋势图')