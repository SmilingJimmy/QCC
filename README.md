# QCC

### 介绍
A  **Python**  based  **Crawler**  programme to get company information from  **QCC**.
Copyright @ _Smiling Jimmy_ 

### 代码文件

-  [qcc_bs4.py](qcc_bs4.py)使用 **BeautifulSoup** 库

   主要整合自网络

-  [qcc_webdriver.py](qcc_webdriver.py)使用selenium的 **webdriver** 库

   主要原创
  
   准确度更高

   应对反爬虫的效果更好

### 安装流程

- 安装 **Python** 或者 **Anaconda** （二选一即可）

  **Python** 不含下述第三方库

  **Anaconda** 包含部分第三方库，如pandas和numpy

- 检查第三方库的安装情况；若有未安装的第三方库，则需要使用 **pip** 安装

  pandas

  numpy

  selenium

  bs4

  requests

  fake_useragent

  fake_useragent第三方库需要在[fake-useragent](https://github.com/hellysmile/fake-useragent)下载 **fake_useragent** 源码安装

  但是，fake_useragent在国内不能正常使用，请在[local_ua](https://github.com/Mehaei/local_ua)下载 **local_ua** 源码并替换fake_useragent的内容

- 下载 **Python** 代码文件到电脑

- 若使用qcc_webdriver.py，需要先安装**谷歌浏览器**，并在[chromedriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)下载谷歌浏览器**第三方驱动和插件**，打开此网站需要 **VPN** 

  如谷歌浏览器版本为87.0.4280.88时，应下载对应版本的[chromedriver](chromedriver.exe)

- 可参考百度或者 **CSDN** 上的安装操作


### 使用说明

- 打开Python代码文件（[qcc_bs4.py](qcc_bs4.py)与[qcc_webdriver.py](qcc_webdriver.py)），修改文件中所有路径为所在电脑的**绝对路径**，注意使用\\\或者/

   要检索公司的 **EXCEL** 文件路径

   爬虫结果输出的 **CSV** 路径

   谷歌浏览器第三方驱动和插件 **EXE** 的路径

- 井号是“ **注释** ”的意思，即该行代码或文字不运行；对于特定行代码，在适当时候可以注释掉（不运行）或取消注释（运行）

- Python代码注重**缩进**，故要避免意外增加或删除空格；一般来说，4个空格代表一层缩进

- 检查用户名、密码、工作簿名称、列名是否正确

- 修改公司的起始索引（含）和结束索引（不含），注意索引从 **0** 开始

- 每次运行文件最好更换爬虫结果输出的 **CSV** 路径，否则同一个 **CSV** 文件中会多次出现标题行

  若不想创建多个 **CSV** 文件，可修改有关输出 **header** 的代码

  [qcc_bs4.py](qcc_bs4.py)： **方法5：整合爬虫方法** 中`df.to_csv(output_path,mode='a+',index=False,header=True,encoding="ANSI")`，header=True改为header=False

  [qcc_webdriver.py](qcc_webdriver.py)： **方法6：把第一个企业的信息输出到CSV文件** 中注释掉`writer.writeheader()`，即`#writer.writeheader()`

- [qcc_webdriver.py](qcc_webdriver.py) **方法1：打开浏览器** 中，`option.add_argument("--headless")`意味着不打开浏览器（在“ **暗地里** ”运行）

  在不熟悉运行步骤或出现报错时，建议注释掉该行代码，即`#option.add_argument("--headless")`，方便直观观察运行情况

  在成熟运行时，为了不影响日常工作，可以保留该行代码，即`option.add_argument("--headless")`

- 在网速较慢，浏览器加载页面较慢时，建议把有关`time.sleep()`中停顿的秒数增大

- 谷歌浏览器占用电脑内存较大，使用webdriver时现需要每搜索 **75** 个公司关闭一次浏览器，即`if round(index%75)==0`；若所在电脑内存消耗较快，后期浏览器页面崩溃，建议将 **75** 改为 **50** 或者 **25** 

- 其余注意事项或释义详见 **Python** 文件中的注释行，此处不再赘述

- 两个代码文件（[qcc_bs4.py](qcc_bs4.py)与[qcc_webdriver.py](qcc_webdriver.py)）互相独立，按需要选择**其中一个**运行即可；以下运行方式选择**其中一个**即可

  在Python**编辑器**（种类繁多，任一即可，安装Python时自带的编辑器也行）中运行代码
  
  在Python代码文件所在文件夹的空白处，按住Shift键点右键，打开**Powershell**窗口，键入`python qcc_bs4.py`或者`python qcc_webdriver.py`，回车即可运行
  
  右键Windows徽标，点击“运行”，输入**cmd**并回车，键入类似于`python E:\钧资料3\大学资料\事务\实习\碧桂园运营实习\供应商信息整理\qcc_bs4.py`或者`python E:\钧资料3\大学资料\事务\实习\碧桂园运营实习\供应商信息整理\qcc_webdriver.py`（带绝对路径），回车即可运行

- 可参考百度或者 **CSDN** 上的操作方法

- Copyright @ _Smiling Jimmy_
