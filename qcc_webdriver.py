#python
#qcc_webdriver.py
#qcc爬虫程序
#by 何智钧
#使用指引：方法1-11如果没有报错的话不需要修改；需要修改的只有文件开头自行输入信息部分！！！

#请自行安装python
#导入第三方库，可能需要pip安装它们
import time
import csv
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from random import choice
#fake_useragent第三方库需要在https://github.com/hellysmile/fake-useragent下载源码安装
#但是，fake_useragent在国内不能正常使用，请在https://github.com/Mehaei/local_ua下载源码并替换fake_useragent的内容
from fake_useragent import UserAgent

#这里请自行输入信息！！！
#企查查的用户名
user_name="13631454966"
#user_name="13129038946"
#企查查的密码
password="bgy2020@"
#password="1q2w3e4r."
#要检索公司的EXCEL文件路径，注意使用\\或者/
company_file_path="E:/钧资料3/大学资料/事务/实习/碧桂园运营实习/供应商信息整理/营销系统常规供应商库.xls"
#company_file_path="E:/钧资料3/大学资料/事务/实习/碧桂园运营实习/供应商信息整理/营销系统战略供应商库.xls"
#要检索公司的EXCEL中工作簿名称
company_file_sheet="Sheet1"
#要检索公司的EXCEL中公司所在列的列名
company_file_column="名称"
#开始搜索的公司索引
begin_index=12730
#结束搜索的公司索引
end_index=15912
#爬虫结果输出的CSV路径，注意使用\\或者/
result_output_path="E:\\钧资料3\\大学资料\\事务\\实习\\碧桂园运营实习\\供应商信息整理\\webdriver_result12730-15912.csv"
#result_output_path="E:\\钧资料3\\大学资料\\事务\\实习\\碧桂园运营实习\\供应商信息整理\\webdriver_result0-933.csv"

#方法1：打开浏览器
def open_browser()->webdriver.Edge:
    #伪造UA
    ua = UserAgent().rget
   #修改请求头
    option = webdriver.ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    option.add_experimental_option('useAutomationExtension', False)
    option.add_argument('--User-Agent={}'.format(ua))
    option.add_argument('--Cookie="zg_did=%7B%22did%22%3A%20%22175967c5e281f9-05201cd90e6fe7-7d677c6e-144000-175967c5e2936a%22%7D; UM_distinctid=17747c7174615a-005eedd542e498-50391c46-144000-17747c717475e9; hasShow=1; _uab_collina=161181696858879448958925; CNZZDATA1254842228=1453378552-1611815288-%7C1611826088; acw_tc=b702c8a116118368249508884e6e7daebca42a35d4f4524c2b2355a258; QCCSESSID=7idbpttfm98na1r88hihqib972; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201611836825549%2C%22updated%22%3A%201611836858753%2C%22info%22%3A%201611816963394%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22cn.bing.com%22%2C%22zs%22%3A%200%2C%22sc%22%3A%200%2C%22cuid%22%3A%20%22c361cd545492bc9b35ec7c9601f0a492%22%7D"')
    option.add_argument('--referer="https://www.qcc.com/"')    
    option.add_argument("--no-sandbox")
    option.add_argument("--disable-dev-usage")
    #option.add_argument("--headless") #无浏览器页面
    option.add_argument("--disable-gpu")
    option.add_argument('log-level=3') #INFO = 0 WARNING = 1 LOG_ERROR = 2 LOG_FATAL = 3
    option.add_argument("--disable-cache") #禁用缓存
    #option.add_argument("--enable-automatic-tab-discarding")
    option.add_argument('--incognito') #无痕隐身模式
    option.add_argument('blink-settings=imagesEnabled=false') #不加载图片, 提升速度
    option.page_load_strategy = 'none' #加载方式，不需要全部加载完成
    desired_capabilities = DesiredCapabilities.CHROME
    desired_capabilities["pageLoadStrategy"] = "none"
    
    #打开浏览器
    #webdriver需要在https://sites.google.com/a/chromium.org/chromedriver/downloads下载谷歌浏览器第三方驱动和插件，打开此网站需要VPN
    browser=webdriver.Chrome("E:/钧资料3/大学资料/事务/实习/碧桂园运营实习/供应商信息整理/chromedriver_win32/chromedriver.exe",options=option)
    #browser.maximize_window()

    #修改navigator为人类
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
        get: () => false
        })
    """
    })

    #打开网页
    browser.get("https://www.qichacha.com/user_login?back=%2F")
    #browser.get("https://www.qcc.com/")
    time.sleep(2)
    return browser

#方法2：从表格中获得公司列表
def load_company_data(file_path:str,sheet:str,column:str)->list:
    print("Loading company data ...")
    companys_data=pd.read_excel(file_path,sheet_name=sheet,header=0,names=None,verbose=True,engine="xlrd")
    companys_whole = np.array(companys_data[column])
    companys=[company for company in companys_whole if company == company]
    return companys    

#方法3：检索公司并回车
def search_company(company:str,browser:webdriver.Chrome)->None:
    wait=WebDriverWait(browser,5)
    time.sleep(1)
   
    input_company=my_search('searchkey',browser)
    input_company.clear()
    input_company.send_keys(company)
    #print("Sending keys ...")
    time.sleep(0.5)
    input_company.send_keys(Keys.ENTER)
    #print("Sending ENTER ...")
    time.sleep(1)

#方法3辅助：等待
def my_search(id:str,browser:webdriver.Chrome):
    error_status=True
    for times in range(6):
        try:
            input_company=browser.find_element_by_id(id)            
        except:
            time.sleep(0.5)
        else:
            error_status=False
            return input_company
    if error_status:
        raise Exception("Failed at search_company()")

#方法4：跳转页面至详情页
def jump2page(browser:webdriver.Chrome)->str:
    time.sleep(1)
    original_window = browser.current_window_handle
    assert len(browser.window_handles) == 1
    wait=WebDriverWait(browser,5)

    my_click('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div/table/tr[1]/td[3]/div/a',browser)
    
    wait.until(EC.number_of_windows_to_be(2))    
    for window_handle in browser.window_handles:
        if window_handle != original_window:
            browser.switch_to.window(window_handle)
            break

    return original_window

#方法4辅助：等待
def my_click(xpath:str,browser:webdriver.Chrome)->None:
    error_status=True
    for times in range(6):
        try:
            company_result=browser.find_element_by_xpath(xpath)
            company_result.click()
        except:
            time.sleep(0.5)
        else:
            error_status=False
            break
    if error_status:
        raise Exception("Failed at jump2page()")

#方法5：从跳转后的页面获取信息
def get_result(company:str,browser:webdriver.Chrome)->dict:
    time.sleep(0.5)
    wait=WebDriverWait(browser,2)

    try:
        input_company=browser.find_element_by_id('headerKey')
        input_company.send_keys(Keys.PAGE_DOWN)
    except:
        pass
    finally:
        time.sleep(0.5)

    information_dict={}
    information_dict["公司"]=company
    information_dict["名称长度"]=len(company)

    #print("Collecting information ...")
    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[2]/div/div/div[2]/a[1]/h2'))
        information_dict["法定代表人"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[2]/div/div/div[2]/a[1]/h2').text
    except:
        information_dict["法定代表人"]=""
    """
    information_dict["法定代表人"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[1]/td[2]/div/div/div[2]/a[1]/h2',browser)
    
    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[4]'))
        information_dict["登记状态"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[4]').text
    except:
        information_dict["登记状态"]=""
    """
    information_dict["登记状态"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[1]/td[4]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[6]'))
        information_dict["成立日期QCC"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[1]/td[6]').text
    except:
        information_dict["成立日期QCC"]=""
    """
    information_dict["成立日期QCC"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[1]/td[6]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[2]'))
        information_dict["注册资本"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[2]').text
    except:
        information_dict["注册资本"]=""
    """
    information_dict["注册资本"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[2]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[4]'))
        information_dict["实缴资本"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[4]').text
    except:
        information_dict["实缴资本"]=""
    """
    information_dict["实缴资本"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[2]/td[4]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[6]'))
        information_dict["核准日期"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[2]/td[6]').text
    except:
        information_dict["核准日期"]=""
    """
    information_dict["核准日期"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[2]/td[6]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[2]'))
        information_dict["统一社会信用代码"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[2]').text
    except:
        information_dict["统一社会信用代码"]=""
    """
    information_dict["统一社会信用代码"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[3]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[4]'))
        information_dict["组织机构代码"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[4]').text
    except:
        information_dict["组织机构代码"]=""
    """
    information_dict["组织机构代码"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[3]/td[4]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[6]'))
        information_dict["工商注册号"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[3]/td[6]').text
    except:
        information_dict["工商注册号"]=""
    """
    information_dict["工商注册号"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[3]/td[6]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[2]'))
        information_dict["企业类型"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[2]').text
    except:
        information_dict["企业类型"]=""
    """
    information_dict["企业类型"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[5]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[4]'))
        information_dict["营业期限"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[4]').text
    except:
        information_dict["营业期限"]=""
    """
    information_dict["营业期限"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[5]/td[4]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[6]'))
        information_dict["登记机关"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[5]/td[6]').text
    except:
        information_dict["登记机关"]=""
    """
    information_dict["登记机关"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[5]/td[6]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[1]/span[2]/span[2]'))
        information_dict["联系电话QCC"]=browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[1]/span[2]/span[2]').text
    except:
        information_dict["联系电话QCC"]=""
    """
    information_dict["联系电话QCC"]=my_wait('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[1]/span[2]/span[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[2]/span[1]/span[2]'))
        information_dict["邮箱"]=browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[2]/span[1]/span[2]').text
    except:
        information_dict["邮箱"]=""
    """
    information_dict["邮箱"]=my_wait('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[2]/span[1]/span[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[3]'))
        information_dict["官网"]=browser.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[3]').text
    except:
        information_dict["官网"]=""
    """
    information_dict["官网"]=my_wait('//*[@id="company-top"]/div[2]/div[2]/div[5]/div[1]/span[3]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[2]'))
        information_dict["人员规模"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[2]').text
    except:
        information_dict["人员规模"]=""
    """
    information_dict["人员规模"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[6]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[4]'))
        information_dict["参保人数"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[4]').text
    except:
        information_dict["参保人数"]=""
    """
    information_dict["参保人数"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[6]/td[4]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[6]'))
        information_dict["省级"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[6]/td[6]').text
    except:
        information_dict["省级"]=""
    """
    information_dict["省级"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[6]/td[6]',browser)

    
    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[8]/td[2]/span/a'))
        information_dict["企业地址"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[8]/td[2]/span/a').text
    except:
        information_dict["企业地址"]=""
    """
    information_dict["企业地址"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[8]/td[2]/span/a',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[9]/td[2]'))
        information_dict["经营范围"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[9]/td[2]').text
    except:
        information_dict["经营范围"]=""
    """
    information_dict["经营范围"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[9]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[7]/td[2]'))
        information_dict["曾用名"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[7]/td[2]').text
    except:
        information_dict["曾用名"]=""
    """
    information_dict["曾用名"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[7]/td[2]',browser)

    """
    try:
        wait.until(lambda browser: browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[7]/td[4]'))
        information_dict["英文名"]=browser.find_element_by_xpath('//*[@id="Cominfo"]/table/tbody/tr[7]/td[4]').text
    except:
        information_dict["英文名"]=""
    """
    information_dict["英文名"]=my_wait('//*[@id="Cominfo"]/table/tbody/tr[7]/td[4]',browser)
        
    return information_dict

#方法5辅助：等待
def my_wait(xpath:str,browser:webdriver.Chrome)->str:
    value=""
    for times in range(2):
        try:
            value=browser.find_element_by_xpath(xpath).text
        except:
            time.sleep(0.2)
        else:
            break
    return value

#方法6：把第一个企业的信息输出到CSV文件
def write_result_firstrow(information_dict:dict,result_output_path:str)->None:
    #print("Writing result in CSV ...")
    fieldname = list(information_dict.keys())
    with open(result_output_path, "a+", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldname)
        writer.writeheader()
        writer.writerow(information_dict)

#方法7：把其余企业的信息输出到CSV文件
def write_result_notfirstrow(information_dict:dict,result_output_path:str)->None:
    #print("Writing result in CSV ...")
    fieldname = list(information_dict.keys())
    with open(result_output_path, "a+", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldname)
        writer.writerow(information_dict)

#方法8：输入用户名和密码，并登录
def input_user_info(user_name:str,password:str,input_browser:webdriver.Chrome)->None:
    browser=input_browser
    while True:
        time.sleep(2)

        #密码登录
        browser.find_element_by_id('normalLogin').click()
        time.sleep(2)

        #用户名
        browser.find_element_by_id('nameNormal').send_keys(user_name)
        print("Sending user name ...")
        time.sleep(2)

        #密码
        browser.find_element_by_id('pwdNormal').send_keys(password)
        print("Sending password ...")
        time.sleep(2)

        #寻找滑块
        button_id_list=['nc_1_n1z','nc_2_n1z','nc_3_n1z','nc_4_n1z','nc_5_n1z','nc_6_n1z','nc_7n1z','nc_8_n1z','nc_9_n1z']
        button_id_error_status=True
        for button_id in button_id_list:
            try:
                button = browser.find_element_by_id(button_id)
            except:
                pass
            else:
                button_id_error_status=False
                break
        if button_id_error_status:
            browser.refresh()
            time.sleep(1)
            continue

        #移动滑块
        """
        ActionChains(browser).click_and_hold(button).perform()
        ActionChains(browser).move_by_offset(xoffset=308, yoffset=0).perform()
        ActionChains(browser).release().perform()
        """
        move_to_gap(button,get_track(308),browser)
        print("Verifying ...")
        time.sleep(4)

        #验证通过
        verify_xpath_list=['//*[@id="nc_1__scale_text"]/span','//*[@id="nc_1__scale_text"]/span/b','//*[@id="nc_2__scale_text"]/span','//*[@id="nc_2__scale_text"]/span/b','//*[@id="nc_3__scale_text"]/span','//*[@id="nc_3__scale_text"]/span/b','//*[@id="nc_4__scale_text"]/span','//*[@id="nc_4__scale_text"]/span/b','//*[@id="nc_5__scale_text"]/span','//*[@id="nc_5__scale_text"]/span/b','//*[@id="nc_6__scale_text"]/span','//*[@id="nc_6__scale_text"]/span/b','//*[@id="nc_7__scale_text"]/span','//*[@id="nc_7__scale_text"]/span/b','//*[@id="nc_8__scale_text"]/span','//*[@id="nc_8__scale_text"]/span/b','//*[@id="nc_9__scale_text"]/span','//*[@id="nc_9__scale_text"]/span/b']
        verify_error_status=True
        for verify_xpath in verify_xpath_list:
            try:            
                assert browser.find_element_by_xpath(verify_xpath).text=="验证通过"
            except:
                pass
            else:
                verify_error_status=False
                break
        if verify_error_status:
            #browser.refresh()
            #time.sleep(2)
            browser=quit_enter(browser)
            continue
        else:
            browser.find_element_by_xpath('//*[@id="user_login_normal"]/button').click()
            print("Logging in ...\n")
            time.sleep(1)
            break
       

#方法8辅助1：位移方程
def get_track(distance:int)->list:
    # 移动轨迹
    track=[]
    # 当前位移
    current=0
    # 减速阈值
    mid=distance*4/5
    # 计算间隔
    t=0.05
    # 初速度
    v=1

    #设置位移时注意随机性
    while current<distance:
        if current<mid:
            # 加速度为随机值
            a=choice(range(4000,7001,1000))
        else:
            # 加速度为随机值
            a=choice([-a for a in range(50,151,50)])
        v0=v
        # 当前速度
        v=v0+a*t
        # 移动距离
        move=v0*t+1/2*a*t*t
        # 当前位移
        current+=move
        # 加入轨迹
        track.append(round(move))
    return track

#方法8辅助2：拖动滑块
def move_to_gap(button,track,browser:webdriver.Chrome)->None:
    ActionChains(browser).click_and_hold(button).perform()
    for x in track:
        ActionChains(browser).move_by_offset(xoffset=x,yoffset=0).perform()
    time.sleep(0.5)
    ActionChains(browser).release().perform()

#方法9：关闭新标签页，切换到原有标签页并后退
def return_back(original_window:str,browser:webdriver.Chrome)->None:
    assert len(browser.window_handles) == 2
    wait=WebDriverWait(browser,5)

    browser.close()
    wait.until(EC.number_of_windows_to_be(1))
    browser.switch_to.window(original_window)
    browser.back()
    #browser.delete_all_cookies()
    time.sleep(1)

#方法10：整合爬虫方法
def integrated_search_functions(user_name:str,password:str,companys:list,result_output_path:str,begin_index:int,end_index:int,input_browser:webdriver.Chrome)->None:
    browser=input_browser
    for index in range(begin_index,end_index):
        try:
            #方法3：检索公司并回车
            search_company(companys[index],browser)
        except:
            try:
                print("{0}-{1}-{2}".format(index,companys[index],"Failed at search_company()"))                
                browser.back()
            except:
                #索引错误直接跳出循环
                print("\nWrong index!")
                break
        else:
            try:
                #方法4：跳转页面至详情页
                original_window=jump2page(browser)
            except:
                print("{0}-{1}-{2}".format(index,companys[index],"Failed at jump2page()"))                
                browser.back()
                #为了防止内存不足，每75个企业要退出1次
                if round(index%75)==0:
                    browser=quit_enter(browser)
                    input_user_info(user_name,password,browser)
            else:
                try:
                    if index==begin_index:
                        #方法5：从跳转后的页面获取信息
                        #方法6：把第一个企业的信息输出到CSV文件
                        write_result_firstrow(get_result(companys[index],browser),result_output_path)
                    else:
                        #方法5：从跳转后的页面获取信息
                        #方法7：把其余企业的信息输出到CSV文件
                        write_result_notfirstrow(get_result(companys[index],browser),result_output_path)
                except:
                    print("{0}-{1}-{2}".format(index,companys[index],"Failed at get_result()"))                    
                    try:
                        #方法9：关闭新标签页，切换到原有标签页并后退
                        return_back(original_window,browser)
                        #为了防止内存不足，每75个企业要退出1次
                        if round(index%75)==0:
                            browser=quit_enter(browser)
                            input_user_info(user_name,password,browser)
                    except:
                        #若无法返回初始状态，只能跳出循环
                        break
                else:
                    try:
                        #方法9：关闭新标签页，切换到原有标签页并后退
                        return_back(original_window,browser)
                        print("{0}-{1}-{2}".format(index,companys[index],"Successful"))
                    except:
                        print("{0}-{1}-{2}".format(index,companys[index],"Failed at return_back()"))
                        #若无法返回初始状态，只能跳出循环
                        break
                    else:
                        #为了防止内存不足，每75个企业要退出1次
                        if round(index%75)==0:
                            browser=quit_enter(browser)
                            input_user_info(user_name,password,browser)

#方法10辅助：退出再进入浏览器
def quit_enter(input_browser:webdriver.Chrome)->webdriver.Chrome:
    input_browser.quit()
    print("Quiting ...")
    time.sleep(8)
    new_browser=open_browser()
    time.sleep(1)
    return new_browser

#方法11：整个程序的方法
def main_function(user_name:str,password:str,company_file_path:str,company_file_sheet:str,company_file_column:str,begin_index:int,end_index:int,result_output_path:str)->None:
    try:
        #方法2：从表格中获得公司列表
        companys=load_company_data(company_file_path,company_file_sheet,company_file_column)
    except:
        print("Unable to load company data!")
    else:
        try:
            #方法1：打开浏览器
            browser=open_browser()
        except:
            print("Unable to open the browser!")                            
        else:
            try:
                #方法8：输入用户名和密码，并登录
                input_user_info(user_name,password,browser) 
            except Exception as e:
                print("Unable to log in!")
                print(e)
                time.sleep(1)
                browser.quit()
            else:
                try:
                    #方法10：整合爬虫方法
                    integrated_search_functions(user_name,password,companys,result_output_path,begin_index,end_index,browser)
                except KeyboardInterrupt:
                    print("\nCtrl-C Interrupt!")
                except Exception as e:
                    print("\n{}\nPlease send the information to 何智钧！".format(e))
                finally:
                    print("\nThanks for using python for searching information ...\nPlease Check the failed companys by hand ...")
                    time.sleep(1)
                    browser.quit()

#寻找index，统计总数
#companys=load_company_data(company_file_path,company_file_sheet,company_file_column)
#print(companys.index("郑州菲梵家居有限公司"))
#print(len(companys))

#程序到这一行现在才正式开始哦！就一行哈哈哈！
#方法11：整个程序的方法
main_function(user_name,password,company_file_path,company_file_sheet,company_file_column,begin_index,end_index,result_output_path)

#如果有任何问题可以去问何智钧