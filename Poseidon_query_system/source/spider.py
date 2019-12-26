# 请求库
import requests
# 时间库
import time
import threading
from bs4 import BeautifulSoup
import xlwt
import re
import copy
import os

def is_number(num):
    # 判断传入的字符串是否为数字
    pattern = re.compile(r'^[-+]?[-0-9]\d*\.\d*|[-+]?\.?[0-9]\d*$')
    result = pattern.match(num)
    if result:
        return True
    else:
        return False

def if_required(course_name):
    # 判断是否为选修课
    if course_name.find("选") == -1:
        # print(course_name + "不是选修课")
        return 1
    else:
        # print(course_name + "是选修课")
        return 0


def login():

    """
    使用部门账号登录教务网
    传出会话窗口,主页地址

    """

    # 构建一个session会话
    session = requests.session()

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36"}
    #url地址
    url_login = 'http://218.25.35.27:8080/(psqmud550o4vblqcvppap055)/default2.aspx'
    # 构建data数据
    data = {
        "__VIEWSTATE": "dDwxODI0OTM5NjI1Ozs+ErNwwEBfve9YGjMA8xEN6zdawEw=",
        "TextBox1": "xinxi",
        "TextBox2": "office326",
        "RadioButtonList1": "部门".encode("gb2312"),
        "Button1": "",
        "lbLanguage": ""
    }


    ## 测试请求是否构建成功
    # print(data)
    url = "http://218.25.35.27:8080/(psqmud550o4vblqcvppap055)/default2.aspx"
    # 发送post请求准备登陆
    html = session.post(url,data=data,headers=headers)
    # 测试是否登录成功
    # print(html)

    # print("-"*20)
    # print("主页登录成功!")
    # print("-"*20)
    html.encoding = "gb2312"
    # 获取查询个人信息的url
    # print(html.url)
    homePage_url = html.url

    html.close()

    # 会话窗口 主页url
    return session,homePage_url

def choose_grade_Page(session,homePage_url,location_url):

    stuInfor_url = location_url + 'cjcx.aspx'


    # 请求报头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36',
        'Referer':location_url + 'bm_main.sapx?xh=xinxi'
    }

    # get请求的数据
    getData = {
        "xh": "xinxi",
        "xm": "信息学院辅导员".encode("gb2312"),
        "gnmkdm": "N120305 "
    }

    # 准备发送get请求
    ##### 注意，以后get请求一定要以params参数的形式发送
    html = session.get(stuInfor_url,params = getData,headers = headers)
    # 测试是否打开了c成绩查询首页
    # print(html.content.decode("gb2312"))

    grade_page_url = html.url

   # 解析出网页中动态变化的__VIEWSTATE
    bs = BeautifulSoup(html.content.decode("gb2312"),'html.parser')
    __VIEWSTATE = bs.find("input",attrs={"type":"hidden","name":"__VIEWSTATE"}).get("value")

    return grade_page_url,__VIEWSTATE

def get_mono_person_grade(session,grade_page_url,__VIEWSTATE,student_num):

    """

    :param grade_page_url: 当前界面的地址
    :param __VIEWSTATE: 从网页动态解析出的参数
    :param student_num: 学生学号
    :return: 学生成绩列表
    """
    # 请求报头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36',
        'Referer':grade_page_url
    }



    data = {
        "__VIEWSTATE": __VIEWSTATE,
        "DropDownList2": "2001-2001",
        "DropDownList3": "a.xh",
        "DropDownList5": "-->请选择学院<--",
        "TextBox1": str(student_num),
        "Button5": "查 询".encode("gb2312")
    }
    person_grade_html = session.post(grade_page_url, data=data, headers=headers)
    # 测试信息是否抓取成功
    # print(person_grade_html.content.decode("gb2312"))
    # print(person_grade_html)

    grade_page_url = person_grade_html.url

    # 解析出网页中动态变化的__VIEWSTATE
    bs = BeautifulSoup(person_grade_html.content.decode("gb2312"), 'html.parser')
    __VIEWSTATE = bs.find("input", attrs={"type": "hidden", "name": "__VIEWSTATE"}).get("value")

    data = {
        "__VIEWSTATE": __VIEWSTATE,
        "DropDownList1": "2001-2001",
        "DropDownList2": "1",
        "DropDownList3": "a.xh",
        "DropDownList4": str(student_num),
        "DropDownList5": "-->请选择学院<--",
        "TextBox1": str(student_num),
        "Button2": "在校学习成绩查询".encode("gb2312")
    }

    # 请求报头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36',
        'Referer':grade_page_url
    }

    person_grade_html = session.post(grade_page_url, data=data, headers=headers)
    soup = BeautifulSoup(person_grade_html.content.decode('gbk'),"html.parser")

    # 测试信息是否抓取成功
    # print(person_grade_html.content.decode("gb2312"))
    # print(person_grade_html)
    # 测试信息是否定位准确
    # print(soup.fieldset.table)

    stu_name = soup.find_all(selected="selected")[1].text
    # print(soup)
    souce_grade = soup.fieldset.table # 原始表格信息
    # 测试是否能够抽离出成绩信息
    # print(soup.fieldset.table.find_all("tr")[1].find_all("td"))

    grade_list = [] # 单科成绩列表

    # 创建表格
    workbook = xlwt.Workbook(encoding="utf-8")
    booksheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok=True)


    i = 0
    j = 0
    for tr in souce_grade.find_all("tr"):
        subject = []
        for td in tr.find_all("td"):
            # 测试是否解析出字符串
            # print(td.contents)

            # 区分字符串和数字
            if td.contents[0].isdigit():# 整数
                booksheet.write(i,j,int(td.contents[0]))
            elif is_number(td.contents[0]):# 浮点数
                booksheet.write(i, j, float(td.contents[0]))
            else:# 字符串
                booksheet.write(i, j, td.contents[0])


            subject.append(td.contents)
            j+=1
        grade_list.append(subject)
        j=0
        i+=1
    # 测试提取姓名学号
    # print(stu_name.replace("|",""))




    # 计算平均绩点
    del grade_list[0]# 删除第一行,因为第一行不是成绩
    # 将列表按照学年拆分
    year_list = []# 存拆分后的列表
    last_year = grade_list[0][0]#上一个列表的年份
    year = [] # 被拆分后的小列表
    for row in grade_list:
        if row[0] != last_year:
            last_year = row[0]
            year_list.append(year)
            year = []
        year.append(row)
    year_list.append(year)

    # 测试是否拆分成功
    # print(year_list)

    # 开始计算

    total_credits = 0#总学分
    Total_grade_point = 0#总绩点

    #调整表格写入位置,保证平均绩点居中
    i += 3
    j = 0
    for year in year_list:
        # print(year)
        for subject in year:
            if if_required(subject[4][0]) and int(subject[14][0]) == 0:
                try:
                    # print(float(subject[6][0]) * float(subject[7][0]))
                    # print(subject[6][0])

                    Total_grade_point += float(subject[6][0]) * float(subject[7][0])
                except:
                    pass
                total_credits += float(subject[6][0])
        booksheet.write(i, j, subject[0][0] + "学年必修课绩点:")
        j += 1
        # print(total_credits)
        if total_credits is not 0:
            booksheet.write(i,j,float(Total_grade_point/total_credits))
        else:
            booksheet.write(i, j, float(0))
        Total_grade_point = 0
        total_credits = 0
        i += 1
        j = 0

    dir = ".\\" + stu_name.replace('|','')[0:8]
    if os.path.isdir(dir):
        dir = dir + "\\"
        workbook.save(dir + stu_name.replace("|","") + ".xls")# 以表格形式保存
    else:
        os.mkdir(dir + "\\")
        dir = dir + "\\"
        workbook.save(dir + stu_name.replace("|", "") + ".xls")  # 以表格形式保存
    return grade_list

def get_mono_person_grade_thread(session,grade_page_url,__VIEWSTATE,student_num_list):
    """
    捕获线程中出现的错误
    """
    for student_num in student_num_list:
        try:
            get_mono_person_grade(session,grade_page_url,__VIEWSTATE,student_num)
            print(str(student_num) + "查询成功QwQ")
        except:
            print(str(student_num) + "查询错误~_~")


def get_stu_num(session,student_information_url,location_url,class_num,__VIEWSTATE):
    """

    :param session: 窗口
    :param student_information_url:  当前页面的url
    :param location_url:
    :param class_num: 要查询的行政班的序号
    :return: 此班级的学号列表
    """
    # 请求报头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36',
        'Referer':student_information_url
    }

    data = {
        "__VIEWSTATE": __VIEWSTATE,
        "DropDownList1": "a.xzd",
        "TextBox1": class_num,
        "Button3": "查  询",


    }
    class_stu_num_html = session.post(student_information_url, data=data, headers=headers)
    # 测试是否请求成功
    # print(person_grade_html.content.decode("gb2312"))
    # print(class_stu_num_html)
    bs = BeautifulSoup(class_stu_num_html.content.decode("gbk"), 'html.parser')
    # print(bs.find_all(id="DropDownList2"))
    # 未经处理的学号信息
    source_num = bs.find_all(id="DropDownList2")

    # 测试正则是否正确
    # print(re.findall(r'value="(.+)">',str(source_num[0])))

    stu_num_list = re.findall(r'value="(.+)">',str(source_num[0]))
    # print(stu_num_list)
    return stu_num_list


def getPersonPage(session,homePage_url,location_url,):
    """
    进入学生信息查询界面
    传出学生信息查询界面的url
    """

    stuInfor_url = location_url + 'xsxx.aspx'

    # 请求报头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.81 Safari/537.36',
        'Referer':location_url + 'bm_main.sapx?xh=xinxi'
    }
    # get请求的数据
    getData = {
        "xh": "xinxi",
        "xm": "信息学院辅导员".encode("gb2312"),
        "gnmkdm": "N120306 "
    }
    # 准备发送get请求
    ##### 注意，以后get请求一定要以params参数的形式发送
    html = session.get(stuInfor_url,params = getData,headers = headers)
    student_information_url = html.url

    # 测试是否打开了学生首页
    # print(html.content.decode("gb2312"))


    # 解析出网页中动态变化的__VIEWSTATE
    bs = BeautifulSoup(html.content.decode("gb2312"),'html.parser')
    __VIEWSTATE = bs.find("input",attrs={"type":"hidden","name":"__VIEWSTATE"}).get("value")
    return student_information_url,__VIEWSTATE

# 爬虫调度函数
def spider():

    #后面多次会用到的一个基础地址
    location_url = 'http://218.25.35.27:8080/(psqmud550o4vblqcvppap055)/'

    # 会话窗口,主页地址
    session,homePage_url = login()
    # 学生信息界面
    student_information_url,__VIEWSTATE_STU_INF = getPersonPage(session,homePage_url,location_url)
    #获取要查询的行政班级的学号
    stu_num_list = get_stu_num(session, student_information_url, location_url, 170306, __VIEWSTATE_STU_INF)
    # print(stu_num_list)
    # 开始查询成绩
    grade_page_url,__VIEWSTATE_grade = choose_grade_Page(session,homePage_url,location_url)


    # for i in stu_num_list:
        # try:
    get_mono_person_grade(session,grade_page_url,__VIEWSTATE_grade,1703070231)
            # print(i + "查询成功QwQ")
        # except:
        #     print(i + "查询错误~_~")

# spider()




