#!/usr/bin/python
from email import header
from httpx import post
from pymysql import NULL
import requests,re,json,xlrd,urllib.request,random,xlrd,xlwt,os,xlsxwriter
from bs4 import BeautifulSoup
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
from random import Random, random
import pandas as pd  # 可独立处理csv，处理excel需配合openpyxl

Cookies=json.load(open('config.json','r'))
Cookie1=Cookies["Cookie1"]
t1=Cookies["t1"]
m1=Cookies["m1"]
s1=Cookies["s1"]
traineeId=Cookies["traineeId"]
Cookie2=Cookies["Cookie2"]
t2=Cookies["t2"]
m2=Cookies["m2"]
s2=Cookies["s2"]

def endiary(diary):
    return(urllib.request.quote(diary))

def enndo():
    template = """今天{}年{}月{}号,现在已经实习了一段时间。今天天气{}，{}。{}
    """
    nian=2022
    month = "4" #月份
    daystart = 1 #这个月开始的时间
    dayend   = 3 #这个月结束的时间+1

    weather = ["非常的晴朗，令我很开心","下雨了，上班有一点不方便","阴，晒衣服不太好晒","晴转多云，气温有点下降","今天有的刮风，但是比较凉爽","受疫情影响现在天气灰蒙蒙的"]
    #天气
    job_content = ["今天在公司里面写代码","今天外出去谈项目","今天和领导汇报工作"]
    #工作内容
    crap = [
        "我深知不能因为自己是一个实习生，就可以随便的做事，相反的对自己的要求应该更加的严格，要努力的工作，不做错任何一件事，因为有可能做错很小的一件事都有可能会影响到公司。所以我很认真的对待这份工作，希望能得到公司的认可，也希望能够继续留任在公司工作。",
        "通过沟通交谈使我对这份工作有了更深的了解，也在更有针对性的学到一些知识。只要我努力肯付出，就一定有收获。",
        "必须在工作中勤于动手慢慢琢磨，不断学习不断积累。遇到不懂的地方，自己先想方设法解决，实在不行可以虚心请教他人，而没有自学能力的人迟早要被企业和社会所淘汰"
        ] #一些废话
    curr_path = os.path.dirname(os.path.abspath(__file__))
    xlspath = os.path.join(curr_path, 'demo.xlsx')
    print(xlspath)

    sheet_name = '日志'

    # 创建 excel 和 表 对象
    wb = xlsxwriter.Workbook(xlspath)
    wbsheet = wb.add_worksheet(sheet_name)

    # 设置内容样式：字体大小, 字体样式，文本居中
    style = wb.add_format({'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'bold': True})

    # 设置表格样式：列宽、行高设置
    wbsheet.set_column(0, 50, 23)  # 0-50列  设置列宽 23
    for nrow in range(10000):
        wbsheet.set_row(nrow, 28)  # 0-999行 设置行高 28
    values=[]
    nrows=0
    ncols=0
    for d in range(daystart,dayend):
        result = template.format(nian,month,d,random.choice(weather),random.choice(job_content),random.choice(crap))
        #print(result)
        sj=str(nian)+"."+str(month)+"."+str(d)
        wbsheet.write(nrows, ncols, sj)
        wbsheet.write(nrows, ncols+1, result)
        nrows=nrows+1
    # 结束
    wb.close()
    df = pd.read_excel(xlspath, sheet_name='日志')                         # 这个会直接默认读取到这个Excel的第一个表单
    print(df)    

def design(sj,nr):
    url="https://www.xybsyw.com/practice/student/blogs/save.action"
    head={
        'Connection':'close',
        'Content-Length':'1797',
        'sec-ch-ua':'" Not A;Brand";v="99", "Chromium";v="99", "Microsoft Edge";v="99"',
        'n':'content,practicePurpose,practiceContent,practiceRequirement,otherRequirement,practiceDescript,securityBook,responsibilities,selfAppraisal,file',
        't':str(t1),
        'sec-ch-ua-mobile':'?0',
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.39',
        'm':str(m1),
        'Accept':'application/json, text/plain, */*',
        's':str(s1),
        'Content-Type':'application/x-www-form-urlencoded',
        'sec-ch-ua-platform':'"macOS"',
        'Origin':'https//www.xybsyw.com',
        'Sec-Fetch-Site':'same-origin',
        'Sec-Fetch-Mode':'cors',
        'Sec-Fetch-Dest':'empty',
        'Referer':'https//www.xybsyw.com/personal/',
        'Accept-Encoding':'gzip, deflate',
        'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Cookie':str(Cookie1)
   }
    da="traineeId="+str(traineeId)+"&title=%E5%AE%9E%E4%B9%A0%E6%97%A5%E5%BF%97&content="+endiary(nr)+"&status=1&visicty=2&type=d&startDate="+sj
    res=requests.post(url=url,headers=head,data=da,verify=False)
    #print(res.text)
    test1=json.loads(res.text)
    if(test1['code']=="200"):
        print("已完成提交本次日报")
        chaxun()
    else:
        print("并未提交成功请联系专业人员")

def chaxun():
    url="https://www.xybsyw.com/practice/student/blogs/loadTaskList.action"
    datas={
        'schoolTermId':'4875',
        'type':'d'
    }
    head={
        'Connection':'close',
        'Content-Length':'24',
        'Pragma':'no-cache',
        'Cache-Control':'no-cache',
        'sec-ch-ua':'" Not A;Brand";v="99", "Chromium";v="99", "Microsoft Edge";v="99"',
        'n':'content,practicePurpose,practiceContent,practiceRequirement,otherRequirement,practiceDescript,securityBook,responsibilities,selfAppraisal,file',
        't':t2,
        'sec-ch-ua-mobile':'?0',
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36 Edg/99.0.1150.39',
        'm':m2,
        'Accept':'application/json, text/plain, */*',
        's':s2,
        'Content-Type':'application/x-www-form-urlencoded',
        'sec-ch-ua-platform':'"macOS"',
        'Origin':'https//www.xybsyw.com',
        'Sec-Fetch-Site':'same-origin',
        'Sec-Fetch-Mode':'cors',
        'Sec-Fetch-Dest':'empty',
        'Referer':'https//www.xybsyw.com/personal/',
        'Accept-Encoding':'gzip, deflate',
        'Accept-Language':'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Cookie':Cookie2
    }
    res=requests.post(url=url,headers=head,data=datas,verify=False)
    div =str(res.text)
    ret_greed= re.findall(r'"blogsCount":"(.*)","canWrite"',div)
    print("恭喜你同学已成功提交了[----->]"+ret_greed[0])

if __name__=="__main__":
    #design(sj,nr)
    #chaxun()
    #生成日记
    enndo()
    workbook = xlrd.open_workbook('demo.xlsx')
    worksheet1 = workbook.sheet_by_name(u'日志')
    tables1=worksheet1.col_values(0, start_rowx=0, end_rowx=None)
    for i in range(len(worksheet1.col_values(1, start_rowx=0, end_rowx=None))):
        sj=worksheet1.col_values(0, start_rowx=i, end_rowx=i+1)
        nr=worksheet1.col_values(1, start_rowx=i, end_rowx=i+1)
        try:    
            if(sj[0]==NULL):
                print("1")
            else:
                design(str(sj[0]),str(nr[0]))
        except:
            print(chaxun())