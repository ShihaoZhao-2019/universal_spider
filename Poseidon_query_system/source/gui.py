from tkinter import *
from PIL import ImageTk,Image
import time
from spider import *

def gettime():
      timestr = time.strftime("%H:%M:%S") # 获取当前的时间并转化为字符串
      lb2.configure(text=timestr)   # 重新设置标签文本
      root.after(1000,gettime) # 每隔1s调用函数 gettime 自身获取时间

def spider_third(str,if_mono=FALSE):
        # 调用爬虫模块
        # 后面多次会用到的一个基础地址
        location_url = 'http://218.25.35.27:8080/(psqmud550o4vblqcvppap055)/'
        # 会话窗口,主页地址
        session, homePage_url = login()
        txt.insert(END, "主页登录成功!\n")

        btn1["text"] = "メ"
        if if_mono:
                grade_page_url, __VIEWSTATE_grade = choose_grade_Page(session, homePage_url, location_url)
                try:
                        get_mono_person_grade(session, grade_page_url, __VIEWSTATE_grade, int(str))
                        try:
                                txt.insert(END,str + "查询成功QwQ\n")
                        except:
                                return 0
                except:
                        try:
                                txt.insert(END,str + "查询错误~_~\n")
                        except:
                                return 0
                try:
                        txt.insert(END, "查询完毕\n")
                except:
                        return 0
                btn1["text"] = "查询"
        else:
                # 学生信息界面
                student_information_url, __VIEWSTATE_STU_INF = getPersonPage(session, homePage_url, location_url)
                # 获取要查询的行政班级的学号
                stu_num_list = get_stu_num(session, student_information_url, location_url, int(str),
                                           __VIEWSTATE_STU_INF)
                # 开始查询成绩
                grade_page_url, __VIEWSTATE_grade = choose_grade_Page(session, homePage_url, location_url)

                for i in stu_num_list:
                        try:
                               get_mono_person_grade(session, grade_page_url, __VIEWSTATE_grade, int(i))
                               try:
                                        txt.insert(END,i + "查询成功QwQ\n")
                               except:
                                       return 0
                        except:
                               try:
                                        txt.insert(END,i + "查询错误~_~\n")
                               except:
                                        return 0
                try:
                        txt.insert(END, "查询完毕\n")
                except:
                         return 0
                btn1["text"] = "查询"






def on_button():

     # 按钮的回调函数
     str = inp.get()
     inp.delete(0, END)  # 清空输入
     txt.delete('1.0','end')

     if(len(str) == 10):
             t = threading.Thread(target=spider_third,args=(str,1,))
             t.start()

     elif(len(str) == 8 or len(str) == 6):
             t = threading.Thread(target=spider_third,args=(str,0,))
             t.start()
     else:
             txt.insert(END, "输入错误请重新输入\n")





root= Tk()
root.title('闪闪star')
# root.geometry('640x480') # 这里的乘号不是 * ，而是小写英文字母 x
root.iconbitmap(r'.\material_library\moon1.ico')

# 设置背景图片
img =Image.open(r'.\material_library\background.jpg')
background_image = ImageTk.PhotoImage(img)
w = background_image.width()
h = background_image.height()
root.geometry('%dx%d+0+0' % (w,h))
root.resizable(0,0)
background_label = Label(root, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# 设置标题
lb = Label(root,text='波塞冬成绩查询系统',
        bg='#FF7F00',
        fg='blue',
        font=('华文新魏',20),
        width=18,
        height=1,
        relief= RAISED)
lb.pack()

# 设置时间
lb2 = Label(root,text='',fg='red',font=("黑体",18))
lb2.pack(side=BOTTOM)
gettime()

# 设置输入提示标签
lb3 =Label(root,text="输入要查询的专业或班级或学号",relief= FLAT,bg='#FFFFFF"',font=("黑体",11),fg='black')
lb3.place(x=10,y=100)
lb4 = Label(root,text="输入:",relief= FLAT,bg='#FFFFFF"',font=("黑体",11),fg='green')
lb4.place(x=10,y=130,)

# 放置输入框
inp = Entry(root)
inp.place(x=60,y=130)

# 放置按钮
btn1 = Button(root, text='查询', command=on_button,font=("黑体",10),bg='#FFFFFF"',fg='#FF00FF')
btn1.place(x=208, y=131, relwidth=0.05, relheight=0.055)

# 放置提示框
txt = Text(root,fg='#00FFFF')
txt.place(x=10,y=170,height=100,width = 230)


root.mainloop()