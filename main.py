from tkinter import *
from tkinter import messagebox
import requests
from PIL import Image
from lxml import etree
import showScore    # 导入显示模块

# 以下是用到的教务处网址，可替换
url1 = 'http://210.44.113.70/loginAction.do'    #登陆地址
url2 = 'http://210.44.113.70/validateCodeAction.do?random=0.3262711948524253'   #验证码地址
# url3 = 'http://210.44.113.70/bxqcjcxAction.do'  #本学期成绩单地址
url4 = 'http://210.44.113.70/gradeLnAllAction.do?type=ln&oper=qbinfo'   # 全部成绩
url5= 'http://210.44.113.70/logout.do'  # 用于退出登录

# 请求头
header = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0'
    }
# 请求数据
data = {
    'zjh':'',
    'mm':'',
    'v_yzm':''
}

# 创建窗体
root = Tk()
root.title("成绩查询系统")    #标题
root.geometry('360x300')        #窗口大小
root.maxsize(360,300)       # 限制窗口大小
root.minsize(360,300)

#定义退出按钮命令
def exit():
    root.quit()
btn1 = Button(root, text='退出系统', command=exit)
btn1.grid(column=0,row=0,columnspan=2)

# 定义改变账号按钮命令
def change():
    zh.delete(0, END)
    keyword.delete(0, END)
    vyzm.delete(0, END)
    brower.post(url5, headers=header, cookies=cookie)

find = Button(root, text='换个账号登录',command=change)
find.grid(column=4,row=0,columnspan=1)

la1 = Label(root,text='查询前请先获取验证码，否则将无响应')
la1.grid(column=2,row=1,pady=2,columnspan=3)

la1 = Label(root, text='学号:')
la1.grid(column=2, row=2)
zh = Entry(root)
zh.grid(column=3, row=2)
# zh.insert(0,'201620') #通过该语句可向输入框中插入字符串

la1 = Label(root)
la1.grid(column=2,row=3,pady=2)

# 用于监控复选框是否选中
showkeyword = IntVar()
# 由复选框的状态决定是否显示密码
def showkeys():
    if(showkeyword.get()==1):
        keyword['show']=''
    else:
        keyword['show']='*'

la2 = Label(root, text='密码:')
la2.grid(column=2, row=4)
keyword = Entry(root,show='*')
keyword.grid(column=3, row=4)
showkey = Checkbutton(root,text='显示密码',variable=showkeyword,onvalue=1,offvalue=0,command=showkeys)
showkey.grid(column=4,row=4)
# keyword.insert(0,'')

cookie={}

# 获取验证码图片并做处理
def abtain_pic():
    global brower
    brower = requests.session()
    yzm = brower.get(url2)
    global cookie
    cookie = brower.cookies.get_dict()
    with open('d:\score\yzm.png','wb')as f:   # 由于pyinstaller打包需要，将此处路径写成绝对地址，下面雷同
        f.write(yzm.content)
    im = Image.open('d:\score\yzm.png')
    im.save("d:\score\yzm.gif", 'gif')
    image = PhotoImage(file='d:\score\yzm.gif')
    yzm_show.configure(image=image)
    yzm_show.image=image

ab = Button(root,text='获取验证码',command=abtain_pic)
ab.grid(column=3,row=5)

lb2 = Label(root, text='验证码:')
lb2.grid(column=2, row=6)
vyzm = Entry(root)
vyzm.grid(column=3, row=6)
image = PhotoImage(file='d:\score\yzm.gif')
yzm_show = Label(root, image=image)
yzm_show.grid(column=4,row=6)

la1 = Label(root)
la1.grid(column=2,row=7,pady=2)

# 获取分数页的源代码
def abtain_url_score():

    response = brower.get(url4,headers=header).text
    return response

# 查询成绩并显示
def search():
    yzm = vyzm.get()
    zjh = zh.get()
    mm = keyword.get()
    if(yzm == ''or zjh == ''or mm == ''):
        messagebox.showwarning('警告','账号、密码、验证码均不能为空')
    else:
        data['zjh'] = zjh
        data['mm'] = mm
        data['v_yzm'] = yzm
        res = brower.post(url1,data=data,headers=header,cookies=cookie).text
        # print(res)
        resp = etree.HTML(res)
        if(resp.xpath('//strong/font')!=[]):
            messagebox.showwarning('提示',resp.xpath('//strong/font')[0].text)
        elif(resp.xpath("//strong")==[] and resp.xpath('//input')==[]):
            messagebox.showinfo('登录成功','加载成绩中')
            html = abtain_url_score()
            showScore.show(html)    # 登录成功后，调用自定义模块的show函数显示成绩单
        else:
            messagebox.showwarning('提示','验证码错误，请重新获取')

find = Button(root, text='查询',command=search)
find.grid(column=3,row=8,columnspan=1)

la1 = Label(root)
la1.grid(column=2,row=9,pady=2)

tip = Label(root,text='说明:此系统仅用于聊大学生查询成绩使用')
tip.grid(column=1,row=10,columnspan=5)


root.mainloop()
