#_*_encoding:utf-8_*_#
import pandas as pd
import numpy as np


#读取店铺数据
def read_shop(shop_xls_path):#输入文件路径
    shop_name=pd.read_excel(shop_xls_path,u'销售，奖金',index_col=0,usecols=[7,8,9,10,11],header=[0])#店铺名字，用来查找poscode
    shop_name = shop_name.index[0][4:-5]

    sell_xls = pd.read_excel(shop_xls_path, u'销售，奖金', index_col=0, usecols=[0, 2, 3,7], skiprows=[0, 1], header=[0])#读取店铺销售月表
    sell_xls=sell_xls.dropna()#舍去空数据
    sell_xls = sell_xls.rename(columns={u"型号（必填15位）": "Referencewithsize", u"数量": "Qty",u"型号":"Reference w/o size"})#重命名，为了做整合
    return shop_name,sell_xls


#读取汇总表
def read_sellout(sellout_xls_path):
    sellout_xls = pd.read_excel(sellout_xls_path, u'database', index_col=None, usecols=np.arange(19))#读取汇总表的database
    sellout_xls = sellout_xls.dropna()#去除na数据
    return sellout_xls


#读取jwl
def read_jwl(sellout_xls_path):
    jwl_xls=pd.read_excel(sellout_xls_path,'Jwl Ref',index_col=None,usecols=[0,3,4,5,6,7,10])
    return jwl_xls


#读取pos,返回code
def read_pos(sellout_xls_path,shop_name):#两个参数，一个是pos表，一个是店铺名
    pos_xls = pd.read_excel(sellout_xls_path, u'pos ranking')
    #pos_code = pos_xls.loc[pos_xls['POSName'] == shop_name].iloc[0, 0]
    pos_information = pos_xls.loc[pos_xls['POSName'] == shop_name]
    return  pos_information


#根据型号寻找
def find_xin(jwl_xls,xinghao):
    xinghao_info= jwl_xls.loc[jwl_xls['ITEM NUMBER W/O SIZE'] == xinghao]
    return xinghao_info


#整合
def integration(shop_xls_path,sellout_xls_path):
    shop_name, sell_xls=read_shop(shop_xls_path)
    sellout_xls=read_sellout(sellout_xls_path)
    pos_information=read_pos(sellout_xls_path, shop_name)

    #加入jwl
    status=[]
    Category=[]
    Style=[]
    Family=[]
    Description=[]
    RSP=[]

    jwl_xls=read_jwl(sellout_xls_path)
    for xinghao in sell_xls['Reference w/o size']:
        status.append(find_xin(jwl_xls,xinghao).iloc[0, 1])
        Category.append(find_xin(jwl_xls, xinghao).iloc[0, 2])
        Style.append(find_xin(jwl_xls, xinghao).iloc[0, 3])
        Family.append(find_xin(jwl_xls, xinghao).iloc[0, 4])
        Description.append(find_xin(jwl_xls, xinghao).iloc[0, 5])
        RSP.append(find_xin(jwl_xls, xinghao).iloc[0, 6])

    sell_xls['18 status'] = status
    sell_xls['Category'] = Category
    sell_xls['Style'] = Style
    sell_xls['Family'] = Family
    sell_xls['Description'] = Description
    sell_xls['RSP'] =  RSP

    #计算总价
    sell_xls['AMOUNT']=sell_xls['Qty']*sell_xls['RSP']

    Referencewithsize_3=[]#截取型号
    for xinghao in sell_xls['Referencewithsize']:
        Referencewithsize_3.append(xinghao[-3:])

    sell_xls['size']=Referencewithsize_3

    #加入poscode
    pos_code_set = []
    for i in range(sell_xls.shape[0]):
        pos_code_set.append(pos_information.iloc[0, 0])

    sell_xls['Pos Code']=pos_code_set

    # 加入poscompany
    pos_company_set = []
    for i in range(sell_xls.shape[0]):
        pos_company_set.append(pos_information.iloc[0, 1])

    sell_xls['CompanyInformation'] = pos_company_set

    # 加入POSName
    pos_name_set = []
    for i in range(sell_xls.shape[0]):
        pos_name_set.append(pos_information.iloc[0, 2])

    sell_xls['POSName'] = pos_name_set

    # 加入Rigion
    pos_rigion_set = []
    for i in range(sell_xls.shape[0]):
        pos_rigion_set.append(pos_information.iloc[0, 4])

    sell_xls['Rigion'] = pos_rigion_set

    # 加入Gucci Sales
    pos_GucciSales_set = []
    for i in range(sell_xls.shape[0]):
        pos_GucciSales_set.append(pos_information.iloc[0, 6])

    sell_xls['Gucci Sales'] = pos_GucciSales_set

    # 加入Xinyu Sales
    pos_Xinyu_set = []
    for i in range(sell_xls.shape[0]):
        pos_Xinyu_set.append(pos_information.iloc[0, 6])

    sell_xls['Xinyu Sales'] = pos_GucciSales_set

    #整合
    pre=pd.concat([sellout_xls,sell_xls],ignore_index=True,join_axes=[sellout_xls.columns])


    return pre,sell_xls.shape[0]


def merge_xls(shop_xls_path,sellout_xls_path):
    pre,rows=integration(shop_xls_path, sellout_xls_path)
    pos_xls = pd.read_excel(sellout_xls_path, u'pos ranking')
    print(pre.iloc[0,2],pos_xls.iloc[0,0])
    after=pd.merge(pre[-rows:],pos_xls,how='left',on='Pos Code',copy=False)
    return after


def check_search(shop_xls_path,sellout_xls_path):
    pre, rows = integration(shop_xls_path, sellout_xls_path)
    pos_xls = pd.read_excel(sellout_xls_path, u'pos ranking')


shop_xls_path=u'C:\\Users\\huang\\Desktop\\test\\邢台.xls'
sellout_xls_path=u'C:\\Users\\huang\\Desktop\\test\\Sell.xlsx'
pre,rows=integration(shop_xls_path,sellout_xls_path)
#after=merge_xls(shop_xls_path,sellout_xls_path)

#'C:\\Users\\huang\\Desktop\\test\\Sell_test.xlsx'
#sheet_name='database'
#index=None
#pre.to_excel()
#pd.ExcelWriter


