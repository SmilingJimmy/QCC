#python
#qcc_bs4.py
#qcc爬虫程序
#by 何智钧
#使用指引：方法1-7如果没有报错的话不需要修改；需要修改的只有文件结尾自行输入信息部分！！！

#请自行安装python
#导入第三方库，可能需要pip安装它们
from bs4 import BeautifulSoup
import requests
import time
import pandas as pd
import numpy as np

#方法1：登录qcc网站
def load_qcc_website(user:str,password:str)->None:
    # 保持会话
    # 新建一个session对象
    sess = requests.session()
    print("Logging in https://www.qcc.com ...")

    # 添加headers
    afterLogin_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36',
                        'Cookie': 'zg_did=%7B%22did%22%3A%20%22175967c5e281f9-05201cd90e6fe7-7d677c6e-144000-175967c5e2936a%22%7D; UM_distinctid=17747c7174615a-005eedd542e498-50391c46-144000-17747c717475e9; hasShow=1; _uab_collina=161181696858879448958925; CNZZDATA1254842228=1453378552-1611815288-%7C1611826088; acw_tc=b702c8a116118368249508884e6e7daebca42a35d4f4524c2b2355a258; QCCSESSID=7idbpttfm98na1r88hihqib972; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201611836825549%2C%22updated%22%3A%201611836858753%2C%22info%22%3A%201611816963394%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22cn.bing.com%22%2C%22zs%22%3A%200%2C%22sc%22%3A%200%2C%22cuid%22%3A%20%22c361cd545492bc9b35ec7c9601f0a492%22%7D',
                        'referer': 'https://www.qcc.com/'}

    # post请求
    login = {'user':user,'password':password}
    proxies = {'http': 'http://120.76.193.51'}
    sess.post('https://www.qcc.com',data=login,headers=afterLogin_headers,proxies=proxies)
    time.sleep(2)

#方法2：获取某公司的全部信息
def get_company_message(company:str)->str:
    # 查询某公司
    search = sess.get('https://www.qcc.com/search?key={}'.format(company),headers=afterLogin_headers,timeout=10)
    search.raise_for_status()
    search.encoding = 'utf-8'
    soup = BeautifulSoup(search.text,features="html.parser")
    href = soup.find_all('a',{'class': 'title'})[0].get('href')
    time.sleep(1)
    # 点击进入第一位搜索结果的新网站，并返回该网址的文本内容
    details = sess.get(href,headers=afterLogin_headers,timeout=10)
    details.raise_for_status()
    details.encoding = 'utf-8' 
    details_soup = BeautifulSoup(details.text,features="html.parser")
    message = details_soup.text
    time.sleep(1)
    return message

#方法3：获取到的文本进行文本特殊化处理，并将其汇总成一个dataframe,方便后面保存为csv
def message_to_df(message:str,company:str)->pd.DataFrame:
    list_companys = []
    Registration_status = []
    Date_of_Establishment = []
    registered_capital = []
    contributed_capital = []
    Approved_date = []
    Unified_social_credit_code = []
    Organization_Code = []
    companyNo = []
    Taxpayer_Identification_Number = []
    sub_Industry = []
    enterprise_type = []
    Business_Term = []
    Registration_Authority = []
    staff_size = []
    Number_of_participants = []
    sub_area = []
    company_adress = []
    Business_Scope = []

    list_companys.append(company)

    try:
        temp_info=message.split('登记状态')[1].split('\n')[1].split('成立日期')[0].replace(' ','')
        if len(temp_info)<50:
            Registration_status.append(temp_info)
        else:
            Registration_status.append(None)
    except:
        Registration_status.append(None)

    try:
        temp_info=message.split('成立日期')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            Date_of_Establishment.append(temp_info)
        else:
            Date_of_Establishment.append(None)
    except:
        Date_of_Establishment.append(None)

    try:
        temp_info=message.split('注册资本为')[1].split('人民币')[0].replace(' ','')
        if len(temp_info)<50:
            registered_capital.append(temp_info)
        else:
            registered_capital.append(None)
    except:
        try:
            temp_info=message.split('注册资本')[1].split('人民币')[0].replace(' ','')
            if len(temp_info)<50:
                registered_capital.append(temp_info)
            else:
                registered_capital.append(None)
        except:
            registered_capital.append(None)

    try:
        temp_info=message.split('实缴资本')[1].split('人民币')[0].replace(' ','')
        if len(temp_info)<50:
            contributed_capital.append(temp_info)
        else:
            contributed_capital.append(None)
    except:
        contributed_capital.append(None)

    try:
        temp_info=message.split('核准日期')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            Approved_date.append(temp_info)
        else:
            Approved_date.append(None)
    except:
        Approved_date.append(None)

    try:
        credit = message.split('统一社会信用代码')[1].split('\n')[1].replace(' ','')
        if len(credit)<50:
            Unified_social_credit_code.append(credit)
        else:
            Unified_social_credit_code.append(None)
    except:
        try:
            credit = message.split('统一社会信用代码')[3].split('\n')[1].replace(' ','')
            if len(credit)<50:
                Unified_social_credit_code.append(credit)
            else:
                Unified_social_credit_code.append(None)
        except:
            Unified_social_credit_code.append(None)
    
    try:
        temp_info=message.split('组织机构代码')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            Organization_Code.append(temp_info)
        else:
            Organization_Code.append(None)
    except:
        Organization_Code.append(None)

    try:
        temp_info=message.split('工商注册号')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            companyNo.append(temp_info)
        else:
            companyNo.append(None)
    except:
        companyNo.append(None)

    try:
        temp_info=message.split('纳税人识别号')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            Taxpayer_Identification_Number.append(temp_info)
        else:
            Taxpayer_Identification_Number.append(None)
    except:
        Taxpayer_Identification_Number.append(None)

    try:
        sub = message.split('所属行业')[1].split('\n')[1].replace(' ','')
        if len(sub)<50:
            sub_Industry.append(sub)
        else:
            sub_Industry.append(None)
    except:
        try:
            sub = message.split('所属行业')[1].split('为')[1].split('，')[0]
            if len(sub)<50:
                sub_Industry.append(sub)
            else:
                sub_Industry.append(None)
        except:
            sub_Industry.append(None)

    try:
        temp_info=message.split('企业类型')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            enterprise_type.append(temp_info)
        else:
            enterprise_type.append(None)
    except:
        enterprise_type.append(None)

    try:
        temp_info=message.split('营业期限')[1].split('登记机关')[0].split('\n')[-1].replace(' ','')
        if len(temp_info)<50:
            Business_Term.append(temp_info)
        else:
            Business_Term.append(None)
    except:
        Business_Term.append(None)

    try:
        temp_info=message.split('登记机关')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            Registration_Authority.append(temp_info)
        else:
            Registration_Authority.append(None)
    except:
        Registration_Authority.append(None)

    try:
        temp_info=message.split('人员规模')[1].split('人')[0].split('\n')[-1].replace(' ','')
        if len(temp_info)<50:
            staff_size.append(temp_info)
        else:
            staff_size.append(None)
    except:
        staff_size.append(None)

    try:
        temp_info=message.split('参保人数')[1].split('所属地区')[0].replace(' ','').split('\n')[2]
        if len(temp_info)<50:
            Number_of_participants.append(temp_info)
        else:
            Number_of_participants.append(None)
    except:
        Number_of_participants.append(None)

    try:
        temp_info=message.split('所属地区')[1].split('\n')[1].replace(' ','')
        if len(temp_info)<50:
            sub_area.append(temp_info)
        else:
            sub_area.append(None)
    except:
        sub_area.append(None)

    try:
        adress = message.split('经营范围')[0].split('企业地址')[1].split('查看地图')[0].split('\n')[2].replace(' ','')
        if len(adress)<50:
            company_adress.append(adress)
        else:
            company_adress.append(None)
    except:
        try:
            adress = message.split('经营范围')[1].split('企业地址')[1].split()[0]
            if len(adress)<50:
                company_adress.append(adress)
            else:
                company_adress.append(None)
        except:
            company_adress.append(None)

    try:
        temp_info=message.split('经营范围')[1].split('\n')[1].replace(' ','')
        Business_Scope.append(temp_info)
    except:
        Business_Scope.append(None)

    df = pd.DataFrame({'公司':company,\
                      '登记状态':Registration_status,\
                      '成立日期':Date_of_Establishment,\
                      '注册资本':registered_capital,\
                      '实缴资本':contributed_capital,\
                      '核准日期':Approved_date,\
                      '统一社会信用代码':Unified_social_credit_code,\
                      '组织机构代码':Organization_Code,\
                      '工商注册号':companyNo,\
                      '纳税人识别号':Taxpayer_Identification_Number,\
                      '所属行业':sub_Industry,\
                      '企业类型':enterprise_type,\
                      '营业期限':Business_Term,\
                      '登记机关':Registration_Authority,\
                      '人员规模':staff_size,\
                      '参保人数':Number_of_participants,\
                      '所属地区':sub_area,\
                      '企业地址':company_adress,\
                      '经营范围':Business_Scope})

    return df

#方法4：读取公司列表
def load_company_data(file_path:str,sheet:str,column:str)->list:
    print("Loading company data ...")
    try:
        companys_data=pd.read_excel(file_path,sheet_name=sheet,header=0,names=None,verbose=True,engine="xlrd")
        companys_whole = np.array(companys_data[column])
        companys=[company for company in companys_whole if company == company]
        time.sleep(2)
        return companys
    except:
        print("Unable to load data!")
        return []

#方法5：整合爬虫方法
def search_company_information(companys:list,output_path:str,begin_index:int,end_index:int)->None:
    print("Begin to search ...")

    #测试网络
    try:
        get_company_message("Testing")
    except:
        time.sleep(5)
        print("Testing the connection first ...")
        print("Below is the searching information, including errors ...\n")

    for index in range(begin_index,end_index):
        time.sleep(1)
        #尝试获取公司信息
        try:
            #方法2：获取某公司的全部信息
            messages = get_company_message(companys[index])
        except:
            try:
                print("{0}-{1}-{2}".format(index,companys[index],"Failed at get_company_message()"))
            except:
                #索引错误直接跳出循环
                print("\nWrong index!")
                break
        else:
            #尝试将公司信息归类
            try:
                #方法3：获取到的文本进行文本特殊化处理，并将其汇总成一个dataframe,方便后面保存为csv
                df = message_to_df(messages,companys[index])
            except:
                print("{0}-{1}-{2}".format(index,companys[index],"Failed at message_to_df()"))
            else:
                #尝试将归类好的信息写入文档
                try:
                    if index==begin_index:
                        df.to_csv(output_path,mode='a+',index=False,header=True,encoding="ANSI")                    
                    else:
                        df.to_csv(output_path,mode='a+',index=False,header=False,encoding="ANSI")
                    print("{0}-{1}-{2}".format(index,companys[index],"Successful"))
                except:
                    print("{0}-{1}-{2}".format(index,companys[index],"Failed at to_csv()"))

#方法6：整个程序的主方法
def main_function(user_name:str,password:str,company_file_path:str,company_file_sheet:str,company_file_column:str,begin_index:int,end_index:int,result_output_path:str)->None:
    try:
        #方法1：登录qcc网站
        load_qcc_website(user_name,password)
        #方法4：读取公司列表
        companys=load_company_data(company_file_path,company_file_sheet,company_file_column)
        #方法5：整合爬虫方法
        search_company_information(companys,result_output_path,begin_index,end_index)
    except KeyboardInterrupt:
        print("\nCtrl-C Interrupt!")
    except:
        print("\nUnknown error!\nPlease send the information to 何智钧！")
    finally:
        print("\nThanks for using python for searching information ...\nPlease Check the failed companys by hand ...")


#这里请自行输入信息！！！
#企查查的用户名
user_name="13631454966"
#企查查的密码
password="bgy2020@"
#要检索公司的EXCEL文件路径，注意使用\\或者/
company_file_path="E:/钧资料3/大学资料/事务/实习/碧桂园运营实习/供应商信息整理/营销系统常规供应商库.xls"
#要检索公司的EXCEL中工作簿名称
company_file_sheet="Sheet1"
#要检索公司的EXCEL中公司所在列的列名
company_file_column="名称"
#开始搜索的公司索引
begin_index=3243
#结束搜索的公司索引
end_index=3300
#爬虫结果输出的CSV路径，注意使用\\或者/
result_output_path="E:\\钧资料3\\大学资料\\事务\\实习\\碧桂园运营实习\\供应商信息整理\\bs4_result3243-5000.csv"

#程序到这一行现在才正式开始哦！就一行哈哈哈！
#方法7：整个程序的方法
main_function(user_name,password,company_file_path,company_file_sheet,company_file_column,begin_index,end_index,result_output_path)

#如果有任何问题可以去问何智钧